
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WiCS Sync')
    .addItem('Sync from export', 'syncFromExport')
    .addItem('Prepare import tab', 'prepareImportTab')
    .addItem('Show help', 'showHelp')
    .addToUi();
}

/** Create / reset the Qualtrics Export tab with short instructions. */
function prepareImportTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.EXPORT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.EXPORT_SHEET_NAME);
  } else {
    sheet.clear();
  }
  sheet
    .getRange(1, 1)
    .setValue(
      'Paste or File → Import your Qualtrics CSV here (replace this text). ' +
        'Then use WiCS Sync → Sync from export. ' +
        'Qualtrics exports usually have 3 header rows — leave them as-is.'
    );
  SpreadsheetApp.getUi().alert(
    'Ready. Import your Qualtrics CSV into the "' +
      CONFIG.EXPORT_SHEET_NAME +
      '" tab, then run Sync from export.'
  );
}

function showHelp() {
  SpreadsheetApp.getUi().alert(
    'WiCS Sync (half-manual)\n\n' +
      '1. In Qualtrics: Data & Analysis → Export → CSV\n' +
      '2. Import/paste into tab: "' +
      CONFIG.EXPORT_SHEET_NAME +
      '"\n' +
      '3. Menu: WiCS Sync → Sync from export\n\n' +
      'Live tab: "' +
      CONFIG.LIVE_SHEET_NAME +
      '"\n' +
      '• Match people by Email (newest submission wins)\n' +
      '• Track submissions by Response ID (no double-import)\n' +
      '• Only fill columns that already exist on the live sheet\n' +
      '• Manual columns (Checked in, Notes, …) are never overwritten'
  );
}

/**
 * Main entry: read Qualtrics Export tab and upsert into the live RSVPs sheet.
 */
function syncFromExport() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getUi().alert('Another sync is already running. Try again in a moment.');
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var exportSheet = ss.getSheetByName(CONFIG.EXPORT_SHEET_NAME);
    if (!exportSheet) {
      throw new Error(
        'Missing tab "' +
          CONFIG.EXPORT_SHEET_NAME +
          '". Run WiCS Sync → Prepare import tab, then paste your CSV there.'
      );
    }

    var records = parseExportSheet_(exportSheet);
    if (!records.length) {
      throw new Error(
        'No response rows found in "' +
          CONFIG.EXPORT_SHEET_NAME +
          '". Import a Qualtrics CSV (with data rows) and try again.'
      );
    }

    var liveSheet = getOrCreateLiveSheet_(ss);
    ensureHelperColumns_(liveSheet);

    var result = upsertRecords_(liveSheet, records);
    var msg =
      'Sync complete.\n\n' +
      'Added: ' +
      result.added +
      '\nUpdated: ' +
      result.updated +
      '\nSkipped (already imported): ' +
      result.skippedSameId +
      '\nSkipped (older than sheet): ' +
      result.skippedOlder +
      '\nSkipped (no email): ' +
      result.skippedNoEmail +
      '\nExport rows read: ' +
      records.length;

    Logger.log(msg);
    SpreadsheetApp.getUi().alert(msg);
  } catch (err) {
    Logger.log('syncFromExport failed: ' + err);
    SpreadsheetApp.getUi().alert('WiCS Sync failed:\n\n' + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// Parse Qualtrics export sheet
// ---------------------------------------------------------------------------

/**
 * Qualtrics CSV layout (typical):
 *  row 1 = short names (ResponseId, Q1, …)
 *  row 2 = human labels (Response ID, First Name, …)
 *  row 3 = {"ImportId":"..."} metadata
 *  row 4+ = data
 *
 * Also accepts a simple 1-header-row sheet.
 */
function parseExportSheet_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (!values.length) {
    return [];
  }

  // Ignore instruction-only first cell
  if (
    values.length === 1 &&
    String(values[0][0]).indexOf('Paste or File') === 0
  ) {
    return [];
  }

  var headerRowIndex = 0; // 0-based index of human-readable headers
  var shortRowIndex = -1;
  var dataStart = 1;

  if (values.length >= 3 && looksLikeImportIdRow_(values[2])) {
    shortRowIndex = 0;
    headerRowIndex = 1;
    dataStart = 3;
  } else if (values.length >= 2 && looksLikeImportIdRow_(values[1])) {
    // Unusual: labels then ImportId
    headerRowIndex = 0;
    dataStart = 2;
  }

  var labels = values[headerRowIndex].map(function (h) {
    return String(h || '').trim();
  });
  var shorts =
    shortRowIndex >= 0
      ? values[shortRowIndex].map(function (h) {
          return String(h || '').trim();
        })
      : labels.slice();

  var responseIdCol = findColumnIndex_(labels, shorts, [
    'Response ID',
    'ResponseId',
    'Responseid'
  ]);
  var emailCol = findColumnIndex_(labels, shorts, [
    'Email',
    'Recipient Email',
    'RecipientEmail'
  ]);
  var submittedCol = findColumnIndex_(labels, shorts, [
    'Recorded Date',
    'RecordedDate',
    'Submitted',
    'End Date',
    'EndDate'
  ]);

  if (responseIdCol < 0) {
    throw new Error(
      'Could not find a "Response ID" column in the export. ' +
        'Use a standard Qualtrics CSV export.'
    );
  }
  if (emailCol < 0) {
    throw new Error(
      'Could not find an "Email" column in the export. ' +
        'The survey needs an Email question (or Recipient Email).'
    );
  }

  var records = [];
  for (var r = dataStart; r < values.length; r++) {
    var row = values[r];
    if (isEmptyRow_(row)) {
      continue;
    }

    var responseId = String(row[responseIdCol] || '').trim();
    var email = normalizeEmail_(row[emailCol]);
    var submittedRaw = submittedCol >= 0 ? row[submittedCol] : '';
    var submitted = parseDateFlexible_(submittedRaw);

    // Index values by original label, short name, and normalized keys for matching
    var byDisplay = {};
    for (var c2 = 0; c2 < labels.length; c2++) {
      var lab = labels[c2];
      if (lab) {
        byDisplay[lab] = row[c2];
        byDisplay[normalizeHeaderKey_(lab)] = row[c2];
      }
      var sh = shorts[c2];
      if (sh) {
        byDisplay[sh] = row[c2];
        byDisplay[normalizeHeaderKey_(sh)] = row[c2];
      }
    }

    records.push({
      responseId: responseId,
      email: email,
      submitted: submitted,
      submittedRaw: submittedRaw,
      byDisplay: byDisplay
    });
  }

  return records;
}

function looksLikeImportIdRow_(row) {
  for (var i = 0; i < Math.min(row.length, 5); i++) {
    var s = String(row[i] || '');
    if (s.indexOf('ImportId') !== -1) {
      return true;
    }
  }
  return false;
}

function findColumnIndex_(labels, shorts, candidates) {
  var i;
  var c;
  for (c = 0; c < candidates.length; c++) {
    var want = normalizeHeaderKey_(candidates[c]);
    for (i = 0; i < labels.length; i++) {
      if (normalizeHeaderKey_(labels[i]) === want) {
        return i;
      }
    }
    for (i = 0; i < shorts.length; i++) {
      if (normalizeHeaderKey_(shorts[i]) === want) {
        return i;
      }
    }
  }
  return -1;
}

function isEmptyRow_(row) {
  for (var i = 0; i < row.length; i++) {
    if (String(row[i] || '').trim() !== '') {
      return false;
    }
  }
  return true;
}

function normalizeEmail_(value) {
  return String(value || '')
    .trim()
    .toLowerCase();
}

function normalizeHeaderKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[_]+/g, ' ')
    .replace(/\s+/g, ' ');
}

function parseDateFlexible_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }
  var s = String(value || '').trim();
  if (!s) {
    return null;
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return d;
  }
  // Qualtrics often: 2025-10-16 20:26:24
  var m = s.match(
    /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/
  );
  if (m) {
    return new Date(
      Number(m[1]),
      Number(m[2]) - 1,
      Number(m[3]),
      Number(m[4]),
      Number(m[5]),
      Number(m[6])
    );
  }
  return null;
}

// ---------------------------------------------------------------------------
// Live sheet helpers
// ---------------------------------------------------------------------------

function getOrCreateLiveSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG.LIVE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LIVE_SHEET_NAME);
    // Sensible starter headers if brand-new
    sheet
      .getRange(1, 1, 1, 6)
      .setValues([
        [
          CONFIG.RESPONSE_ID_HEADER,
          CONFIG.SUBMITTED_HEADER,
          'First Name',
          'Last Name',
          CONFIG.EMAIL_HEADER,
          'Notes'
        ]
      ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** Ensure Response ID / Submitted / Email columns exist (append if missing). */
function ensureHelperColumns_(sheet) {
  var headers = getHeaders_(sheet);
  var needed = [
    CONFIG.RESPONSE_ID_HEADER,
    CONFIG.SUBMITTED_HEADER,
    CONFIG.EMAIL_HEADER
  ];
  var toAdd = [];
  needed.forEach(function (name) {
    if (findHeaderIndex_(headers, name) < 0) {
      toAdd.push(name);
    }
  });
  if (toAdd.length) {
    var startCol = headers.length + 1;
    if (headers.length === 0 || (headers.length === 1 && headers[0] === '')) {
      sheet.getRange(1, 1, 1, toAdd.length).setValues([toAdd]);
      sheet.setFrozenRows(1);
    } else {
      sheet.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
    }
  }
}

function getHeaders_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    return [];
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h || '').trim();
  });
  while (headers.length && headers[headers.length - 1] === '') {
    headers.pop();
  }
  return headers;
}

function findHeaderIndex_(headers, name) {
  var want = normalizeHeaderKey_(name);
  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeaderKey_(headers[i]) === want) {
      return i;
    }
  }
  // aliases
  var aliases = HEADER_ALIASES_[want] || [];
  for (var a = 0; a < aliases.length; a++) {
    for (var j = 0; j < headers.length; j++) {
      if (normalizeHeaderKey_(headers[j]) === aliases[a]) {
        return j;
      }
    }
  }
  return -1;
}

var HEADER_ALIASES_ = {
  'response id': ['responseid', 'response id'],
  submitted: ['recorded date', 'recordeddate', 'end date', 'enddate'],
  email: ['recipient email', 'recipientemail']
};

// ---------------------------------------------------------------------------
// Upsert: Response ID skip + email newest-wins
// ---------------------------------------------------------------------------

function upsertRecords_(liveSheet, records) {
  var headers = getHeaders_(liveSheet);
  var idCol = findHeaderIndex_(headers, CONFIG.RESPONSE_ID_HEADER);
  var emailCol = findHeaderIndex_(headers, CONFIG.EMAIL_HEADER);
  var submittedCol = findHeaderIndex_(headers, CONFIG.SUBMITTED_HEADER);

  if (idCol < 0 || emailCol < 0) {
    throw new Error(
      'Live sheet needs "' +
        CONFIG.RESPONSE_ID_HEADER +
        '" and "' +
        CONFIG.EMAIL_HEADER +
        '" columns.'
    );
  }

  var lastRow = liveSheet.getLastRow();
  var width = headers.length;
  var data = [];
  if (lastRow >= 2) {
    data = liveSheet.getRange(2, 1, lastRow - 1, width).getValues();
  }

  // Maps for lookup
  var rowIndexByResponseId = {}; // responseId -> 0-based index in data
  var rowIndexByEmail = {}; // email -> 0-based index in data

  for (var i = 0; i < data.length; i++) {
    var rid = String(data[i][idCol] || '').trim();
    var em = normalizeEmail_(data[i][emailCol]);
    if (rid) {
      rowIndexByResponseId[rid] = i;
    }
    if (em) {
      rowIndexByEmail[em] = i;
    }
  }

  var added = 0;
  var updated = 0;
  var skippedSameId = 0;
  var skippedOlder = 0;
  var skippedNoEmail = 0;
  var rowsToAppend = [];

  // Process export records newest-first so first write for an email is the latest
  var sorted = records.slice().sort(function (a, b) {
    var ta = a.submitted ? a.submitted.getTime() : 0;
    var tb = b.submitted ? b.submitted.getTime() : 0;
    return tb - ta;
  });

  var seenEmailInThisRun = {};

  sorted.forEach(function (rec) {
    if (!rec.responseId) {
      return;
    }

    // Already imported this exact Qualtrics submission
    if (rowIndexByResponseId[rec.responseId] !== undefined) {
      skippedSameId++;
      return;
    }

    if (!rec.email) {
      skippedNoEmail++;
      return;
    }

    // Within this export, only keep the newest row per email (sorted newest-first)
    if (seenEmailInThisRun[rec.email]) {
      skippedOlder++;
      return;
    }
    seenEmailInThisRun[rec.email] = true;

    var existingIdx = rowIndexByEmail[rec.email];

    if (existingIdx === undefined) {
      // New person → append
      var newRow = buildRowValues_(headers, rec, null);
      rowsToAppend.push(newRow);
      // Track as if already in sheet so later logic is consistent
      var fakeIdx = data.length + rowsToAppend.length - 1;
      // We'll remap after append; for same-run duplicates we already gated by seenEmailInThisRun
      rowIndexByEmail[rec.email] = 'pending:' + (rowsToAppend.length - 1);
      rowIndexByResponseId[rec.responseId] = 'pending';
      added++;
      return;
    }

    if (typeof existingIdx === 'string' && existingIdx.indexOf('pending:') === 0) {
      // Already queued an append for this email in this run (newer one) — skip
      skippedOlder++;
      return;
    }

    // Existing person — update only if this submission is newer (or sheet has no date)
    var existingSubmitted =
      submittedCol >= 0 ? parseDateFlexible_(data[existingIdx][submittedCol]) : null;
    if (
      existingSubmitted &&
      rec.submitted &&
      rec.submitted.getTime() <= existingSubmitted.getTime()
    ) {
      skippedOlder++;
      return;
    }

    var updatedRow = buildRowValues_(headers, rec, data[existingIdx]);
    data[existingIdx] = updatedRow;
    rowIndexByResponseId[rec.responseId] = existingIdx;
    // Write this row immediately (safer than batching mixed updates)
    liveSheet
      .getRange(existingIdx + 2, 1, 1, width)
      .setValues([updatedRow]);
    updated++;
  });

  if (rowsToAppend.length) {
    var startRow = liveSheet.getLastRow() + 1;
    liveSheet
      .getRange(startRow, 1, rowsToAppend.length, width)
      .setValues(rowsToAppend);
  }

  return {
    added: added,
    updated: updated,
    skippedSameId: skippedSameId,
    skippedOlder: skippedOlder,
    skippedNoEmail: skippedNoEmail
  };
}

/**
 * Build a full-width row for the live sheet.
 * existingRow: if provided, copy through any columns we shouldn't overwrite (manual).
 */
function buildRowValues_(headers, rec, existingRow) {
  var row = [];
  for (var c = 0; c < headers.length; c++) {
    var header = headers[c];
    var key = normalizeHeaderKey_(header);

    // Always set helper fields from the record
    if (key === normalizeHeaderKey_(CONFIG.RESPONSE_ID_HEADER)) {
      row.push(rec.responseId);
      continue;
    }
    if (key === normalizeHeaderKey_(CONFIG.SUBMITTED_HEADER)) {
      row.push(rec.submitted || rec.submittedRaw || '');
      continue;
    }
    if (key === normalizeHeaderKey_(CONFIG.EMAIL_HEADER)) {
      row.push(rec.email);
      continue;
    }

    var exportVal = lookupExportValue_(rec, header);
    if (exportVal !== undefined) {
      row.push(exportVal);
    } else if (existingRow) {
      // Manual / unmatched column — preserve
      row.push(existingRow[c]);
    } else {
      row.push('');
    }
  }
  return row;
}

/** Find a value in the export record for a live-sheet header name. */
function lookupExportValue_(rec, liveHeader) {
  var want = normalizeHeaderKey_(liveHeader);
  if (rec.byDisplay[liveHeader] !== undefined) {
    return rec.byDisplay[liveHeader];
  }
  if (rec.byDisplay[want] !== undefined) {
    return rec.byDisplay[want];
  }

  // Alias map for common Qualtrics vs sheet naming
  var aliases = {
    'first name': ['first name', 'q1', 'recipient first name'],
    'last name': ['last name', 'q3', 'recipient last name'],
    email: ['email', 'q10', 'recipient email'],
    submitted: ['recorded date', 'end date', 'submitted'],
    'response id': ['response id', 'responseid']
  };
  var list = aliases[want] || [want];
  for (var i = 0; i < list.length; i++) {
    if (rec.byDisplay[list[i]] !== undefined) {
      return rec.byDisplay[list[i]];
    }
  }

  // Last resort: scan export labels for normalized equality
  for (var k in rec.byDisplay) {
    if (normalizeHeaderKey_(k) === want) {
      return rec.byDisplay[k];
    }
  }
  return undefined;
}

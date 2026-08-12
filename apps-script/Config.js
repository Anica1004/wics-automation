/**
 * Half-manual Qualtrics → Google Sheets sync (no Qualtrics API).
 *
 * Flow:
 *  1. Export CSV from Qualtrics
 *  2. Import / paste it into the "Qualtrics Export" tab
 *  3. WiCS Sync → Sync from export
 *
 * Behavior:
 *  - One row per email on the live sheet (newest submission wins)
 *  - Response ID uniquely identifies each Qualtrics submission (skip if already imported)
 *  - Only writes columns that already exist on the live sheet (matched by header name)
 *  - Never overwrites unmatched / manual columns (Checked in, Notes, …)
 */
var CONFIG = {
  /** Live roster tab — keep your existing headers here. */
  LIVE_SHEET_NAME: 'RSVPs',
  /** Paste / File→Import the Qualtrics CSV into this tab. */
  EXPORT_SHEET_NAME: 'Qualtrics Export',
  /** Canonical helper headers (created on the live sheet if missing). */
  RESPONSE_ID_HEADER: 'Response ID',
  SUBMITTED_HEADER: 'Submitted',
  EMAIL_HEADER: 'Email',
  /**
   * Live-sheet columns the script is allowed to write when they match the export.
   * Leave empty [] to allow any column that matches an export header (recommended).
   * Manual columns simply won't match export headers, so they stay untouched.
   */
  WRITEABLE_HEADERS: []
};

# WiCS Qualtrics → Google Sheets sync (Plan B — half-manual)

UBC Qualtrics blocks API access for typical student accounts (`403 AuthZ`), so this sync **does not call Qualtrics**. You export a CSV, paste it into the sheet, and Apps Script merges it.

## How it works

1. Export **CSV** from Qualtrics (Data & Analysis → Export).
2. Paste or **File → Import** it into the **`Qualtrics Export`** tab.
3. **WiCS Sync → Sync from export**.

The script then updates the **`RSVPs`** tab:

| Rule | Behavior |
|------|----------|
| **Email** | Identifies a person → one row per email |
| **Response ID** | Identifies a Qualtrics submission → skip if already imported |
| **Newest wins** | If that email already exists and the export row is newer → update Qualtrics columns only |
| **Live headers** | Only fills columns that already exist on `RSVPs` (matched by name) |
| **Manual columns** | e.g. Checked in, Notes — never overwritten |

## Setup

1. Open your WiCS Google Sheet (or create one).
2. **Extensions → Apps Script**.
3. Replace default code with:
   - [`apps-script/Config.js`](apps-script/Config.js) → file `Config.gs`
   - [`apps-script/Code.js`](apps-script/Code.js) → file `Code.gs`
4. Save → run **`onOpen`** once (or reload the sheet) to get the **WiCS Sync** menu.
5. **WiCS Sync → Prepare import tab** (creates `Qualtrics Export`).
6. Make sure the live tab is named **`RSVPs`** (or change `LIVE_SHEET_NAME` in `Config.gs`).
7. Put your column headers on `RSVPs` (First Name, Last Name, Email, Notes, …). The script will add **Response ID** / **Submitted** / **Email** if they’re missing.

### Day-to-day

1. Export CSV from Qualtrics.  
2. Import into **`Qualtrics Export`** (replace old contents).  
3. **WiCS Sync → Sync from export**.  
4. Read the summary (Added / Updated / Skipped).

## Config (`Config.gs`)

```js
LIVE_SHEET_NAME: 'RSVPs',
EXPORT_SHEET_NAME: 'Qualtrics Export',
RESPONSE_ID_HEADER: 'Response ID',
SUBMITTED_HEADER: 'Submitted',
EMAIL_HEADER: 'Email',
```

Rename tabs/headers there if your sheet uses different names — keep **Email** and **Response ID** available on the live sheet.

## Notes

- Prefer **CSV** over Excel (Qualtrics’ 3 header rows are handled automatically).
- Blank emails are skipped (script can’t match a person).
- Re-importing the same export is safe: already-seen Response IDs are skipped.
- The old API-based sync is retired; no token needed.
# wics-automation

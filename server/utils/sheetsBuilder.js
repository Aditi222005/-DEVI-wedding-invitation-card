const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');

// ─── CONFIG ───────────────────────────────────────────────────────────────────
const SPREADSHEET_ID = '1CTKy_mqIspnI9dcUhxJ9uHEgAwenw2e9M4_f94pOW6s';
const SHEET_NAME = 'DEVI-formal-invitation-data';          // Tab name in the Google Sheet
const CREDS_PATH = path.join(__dirname, '..', 'google-credentials.json');

// ─── AUTH ─────────────────────────────────────────────────────────────────────
function getAuth() {
  if (!fs.existsSync(CREDS_PATH)) {
    throw new Error(
      'google-credentials.json not found in server/. ' +
      'Please follow the setup guide to create a service account and place the key file there.'
    );
  }

  const auth = new google.auth.GoogleAuth({
    keyFile: CREDS_PATH,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return auth;
}

// ─── ENSURE HEADER ROW EXISTS ─────────────────────────────────────────────────
async function ensureHeader(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A1:H1`,
  });

  const firstRow = res.data.values?.[0] ?? [];
  if (firstRow.length === 0) {
    // Sheet is empty — write the header
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [[
          '#', 'Name', 'Guests', 'Attending', 'Submitted At (IST)', 'Device ID', 'IP Address', 'User Agent'
        ]],
      },
    });
  }
}

// ─── APPLY HEADER FORMATTING (bold + background) ──────────────────────────────
async function formatHeader(sheetsApi, sheetId) {
  await sheetsApi.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 8 },
            cell: {
              userEnteredFormat: {
                backgroundColor: { red: 0.42, green: 0.10, blue: 0.10 }, // maroon
                textFormat: {
                  bold: true,
                  foregroundColor: { red: 0.94, green: 0.88, blue: 0.66 }, // gold-pale
                  fontSize: 11,
                },
                horizontalAlignment: 'CENTER',
              },
            },
            fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)',
          },
        },
        {
          updateSheetProperties: {
            properties: { sheetId, gridProperties: { frozenRowCount: 1 } },
            fields: 'gridProperties.frozenRowCount',
          },
        },
      ],
    },
  });
}

// ─── GET SHEET ID (numeric) FOR FORMATTING ────────────────────────────────────
async function getSheetId(sheetsApi) {
  const meta = await sheetsApi.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheet = meta.data.sheets.find(s => s.properties.title === SHEET_NAME);
  return sheet ? sheet.properties.sheetId : 0;
}

// ─── APPEND A ROW ─────────────────────────────────────────────────────────────
/**
 * Appends one RSVP row to the Google Sheet.
 * @param {object} rsvp  - Mongoose RSVP document
 * @param {number} rowNum - The row number (used as serial #)
 */
async function appendRSVPRow(rsvp, rowNum) {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // First time setup
  await ensureHeader(sheets);

  const ist = new Date(rsvp.submittedAt).toLocaleString('en-IN', {
    timeZone: 'Asia/Kolkata',
    day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit', hour12: true,
  });

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:H`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: [[
        rowNum,
        rsvp.name,
        rsvp.guestCount,
        rsvp.attending ? '✓ Yes' : '✗ No',
        ist,
        rsvp.deviceId,
        rsvp.ipAddress || '—',
        (rsvp.userAgent || '—').substring(0, 80),
      ]],
    },
  });

  console.log(`📊  Google Sheet updated — row ${rowNum}: ${rsvp.name}`);
}

// ─── FULL REBUILD (used by /download or on demand) ────────────────────────────
/**
 * Clears the sheet and rewrites all RSVP rows from the DB.
 */
async function rebuildSheet(rsvps) {
  const auth = getAuth();
  const sheets = google.sheets({ version: 'v4', auth });

  // Clear existing data
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:H`,
  });

  const header = [['#', 'Name', 'Guests', 'Attending', 'Submitted At (IST)', 'Device ID', 'IP Address', 'User Agent']];
  const rows = rsvps.map((r, i) => {
    const ist = new Date(r.submittedAt).toLocaleString('en-IN', {
      timeZone: 'Asia/Kolkata',
      day: '2-digit', month: 'short', year: 'numeric',
      hour: '2-digit', minute: '2-digit', hour12: true,
    });
    return [
      i + 1,
      r.name,
      r.guestCount,
      r.attending ? '✓ Yes' : '✗ No',
      ist,
      r.deviceId,
      r.ipAddress || '—',
      (r.userAgent || '—').substring(0, 80),
    ];
  });

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A1`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [...header, ...rows] },
  });

  // Apply formatting
  try {
    const sheetId = await getSheetId(sheets);
    await formatHeader(sheets, sheetId);
  } catch (_) { /* formatting is cosmetic — ignore errors */ }

  console.log(`📊  Google Sheet fully rebuilt — ${rsvps.length} entries`);
}

module.exports = { appendRSVPRow, rebuildSheet };

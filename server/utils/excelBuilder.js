const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const EXCEL_PATH = path.join(__dirname, '..', 'data', 'rsvp_list.xlsx');

// Ensure data directory exists
function ensureDir() {
  const dir = path.dirname(EXCEL_PATH);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

/**
 * Rebuild the Excel file from an array of RSVP documents.
 * Called after every new RSVP insert.
 */
async function rebuildExcel(rsvps) {
  ensureDir();

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Viraj & Devaki Wedding System';
  wb.created = new Date();

  const ws = wb.addWorksheet('RSVP List', {
    pageSetup: { fitToPage: true, fitToWidth: 1 },
  });

  // ── Header styling ──────────────────────────────────────────────────────────
  const headerFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF6B1A1A' }, // maroon
  };
  const headerFont = {
    bold: true,
    color: { argb: 'FFF0E0A8' }, // gold-pale
    name: 'Calibri',
    size: 12,
  };
  const centerAlign = { horizontal: 'center', vertical: 'middle' };

  // Column definitions
  ws.columns = [
    { header: '#',           key: 'sno',         width: 6  },
    { header: 'Name',        key: 'name',        width: 28 },
    { header: 'Guests',      key: 'guestCount',  width: 10 },
    { header: 'Attending',   key: 'attending',   width: 14 },
    { header: 'Submitted At',key: 'submittedAt', width: 24 },
    { header: 'Device ID',   key: 'deviceId',    width: 38 },
    { header: 'IP Address',  key: 'ipAddress',   width: 18 },
  ];

  // Style header row
  const headerRow = ws.getRow(1);
  headerRow.height = 22;
  headerRow.eachCell(cell => {
    cell.fill   = headerFill;
    cell.font   = headerFont;
    cell.alignment = centerAlign;
    cell.border = {
      bottom: { style: 'medium', color: { argb: 'FFB8962E' } },
    };
  });

  // ── Data rows ───────────────────────────────────────────────────────────────
  const acceptFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
  const declineFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4EC' } };

  rsvps.forEach((r, idx) => {
    const row = ws.addRow({
      sno:         idx + 1,
      name:        r.name,
      guestCount:  r.guestCount,
      attending:   r.attending ? 'Yes ✓' : 'No ✗',
      submittedAt: new Date(r.submittedAt).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      deviceId:    r.deviceId,
      ipAddress:   r.ipAddress || '—',
    });

    // Alternate + attending colour
    row.eachCell(cell => {
      cell.fill = r.attending ? acceptFill : declineFill;
      cell.alignment = { vertical: 'middle' };
    });
    row.getCell('sno').alignment = centerAlign;
    row.getCell('guestCount').alignment = centerAlign;
    row.getCell('attending').alignment = centerAlign;
  });

  // ── Summary row ─────────────────────────────────────────────────────────────
  ws.addRow([]);
  const totalAttending = rsvps.filter(r => r.attending).reduce((s, r) => s + r.guestCount, 0);
  const totalDeclined  = rsvps.filter(r => !r.attending).length;

  const summaryRow = ws.addRow([
    '',
    `Total RSVPs: ${rsvps.length}`,
    `Attending guests: ${totalAttending}`,
    `Declined: ${totalDeclined}`,
  ]);
  summaryRow.font = { bold: true, name: 'Calibri', size: 11 };
  summaryRow.getCell(2).fill = headerFill;
  summaryRow.getCell(2).font = { ...summaryRow.font, color: { argb: 'FFF0E0A8' } };

  // Auto-filter on header
  ws.autoFilter = { from: 'A1', to: 'G1' };

  await wb.xlsx.writeFile(EXCEL_PATH);
  return EXCEL_PATH;
}

module.exports = { rebuildExcel, EXCEL_PATH };

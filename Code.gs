// Event check-in on Google Sheets. Reuse for any event: File > Make a copy of this Sheet.
const TZ = 'Asia/Kathmandu';   // Nepal time

// Column order lives in ONE place. Header and column numbers are both built from this list.
// Keep IN, TIME, BY next to each other (a scan writes those three cells together).
const COLS = [
  ['SERIAL', 'SN'], ['NAME', 'Name'], ['ORG', 'Organization'], ['DESIG', 'Designation'], ['PHONE', 'Phone'],
  ['EMAIL', 'Email'], ['OTHER', 'Other'], ['QRPDF', 'QR PDF'],
  ['IN', 'Checked in'], ['TIME', 'Check-in time (Nepal)'], ['BY', 'Scanned by'],
  ['INV', 'Invoice PDF'], ['CODE', 'Ticket code'], ['QR', 'QR code (download link)']
];

const HEAD = COLS.map(c => c[1]);
const C = {};
COLS.forEach((c, i) => C[c[0]] = i + 1);

function L_(n) {
  let t = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    t = String.fromCharCode(65 + m) + t;
    n = Math.floor((n - 1) / 26);
  }
  return t;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Check-in')
    .addItem('1. Set up sheets', 'setup')
    .addItem('2. Generate QR codes + PDFs', 'generateQR')
    .addItem('3. Create invoice PDFs', 'createInvoices')
    .addToUi();
}

function sheet_() {
  return SpreadsheetApp.getActive().getSheetByName('Attendees');
}

function nowNepal_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd hh:mm:ss a');
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Attendees') || ss.insertSheet('Attendees');
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#e8ebf2');
  sh.setFrozenRows(1);
  const rows = sh.getMaxRows() - 1;
  sh.getRange(2, C.TIME, rows, 1).setNumberFormat('@');          // keep time as exact text
  sh.setColumnWidth(C.QR, 170);
  sh.setColumnWidth(C.TIME, 190);

  // Green when Checked in = YES (strong green on the YES cell, light green on the row)
  const yes = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + L_(C.IN) + '2="YES"')
    .setBackground('#1f9d55')
    .setFontColor('#ffffff')
    .setBold(true)
    .setRanges([sh.getRange(2, C.IN, rows, 1)])
    .build();
  const row = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + L_(C.IN) + '2="YES"')
    .setBackground('#d5efdf')
    .setRanges([sh.getRange(2, 1, rows, HEAD.length)])
    .build();
  sh.setConditionalFormatRules([yes, row]);
  liveSheet_();

  const st = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  if (st.getLastRow() === 0) {
    st.getRange(1, 1, 11, 3).setValues([
      ['Setting', 'Value', 'Note'],
      ['staffKey', Utilities.getUuid().slice(0, 8), 'Password the scanner asks for'],
      ['orgName', 'My Organization', 'Shown on invoices and QR PDFs'],
      ['eventName', 'My Event', 'Shown on invoices and QR PDFs'],
      ['eventDate', '', 'Shown on invoices and QR PDFs'],
      ['venue', '', 'Shown on invoices and QR PDFs'],
      ['address', '', 'Organization address (invoice)'],
      ['contact', '', 'Phone / email on invoice'],
      ['panVat', '', 'PAN / VAT number (optional)'],
      ['logoDriveFileId', '', 'Paste full Google Drive link. Set sharing to "Anyone with the link".'],
      ['invoiceNote', 'Thank you for registering.', 'Invoice footer text']
    ]);
    st.getRange(1, 1, 1, 3).setFontWeight('bold');
    st.setColumnWidth(2, 260);
    st.setColumnWidth(3, 420);
  }
  SpreadsheetApp.getUi().alert('Ready. Add attendees from row 2 (Name is required), edit the Settings sheet, then run "Generate QR codes + PDFs".');
}

function settings_() {
  const o = {};
  SpreadsheetApp.getActive().getSheetByName('Settings').getDataRange().getValues().slice(1)
    .forEach(r => o[r[0]] = r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : r[1]);
  return o;
}

// SPEED: staff key from fast cache instead of reading the Settings sheet on every request.
// If you change staffKey in Settings, the new key works within 5 minutes.
function staffKey_() {
  const cache = CacheService.getScriptCache();
  let k = cache.get('staffKey');
  if (k === null) {
    k = String(settings_().staffKey);
    cache.put('staffKey', k, 300);
  }
  return k;
}

// QR text = Name | Designation | Code. A normal phone camera shows name and designation; the scanner uses the code.
function qrData_(r) {
  return r[C.NAME - 1] + ' | ' + r[C.DESIG - 1] + ' | ' + r[C.CODE - 1];
}

// One step: serial numbers, ticket codes, QR download links AND a PDF (with all details) for every row that has a name.
function generateQR() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return SpreadsheetApp.getUi().alert('Add attendees first (Name column).');
  const v = sh.getRange(2, 1, n, HEAD.length).getValues();
  const serial = [], code = [], qr = [];

  v.forEach((r, i) => {
    const row = i + 2;
    if (!r[C.NAME - 1]) {
      serial.push(['']);
      code.push(['']);
      qr.push(['']);
      return;
    }
    serial.push([r[C.SERIAL - 1] || i + 1]);

    // Only generate a new code if one doesn't exist yet
    const existingCode = r[C.CODE - 1];
    const newCode = existingCode || Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    code.push([newCode]);

    qr.push(['=HYPERLINK("https://api.qrserver.com/v1/create-qr-code/?size=400x400&download=1&data="&ENCODEURL(' + L_(C.NAME) + row + '&" | "&' + L_(C.DESIG) + row + '&" | "&' + L_(C.CODE) + row + '),"Download QR")']);
  });

  sh.getRange(2, C.SERIAL, n, 1).setValues(serial);
  sh.getRange(2, C.CODE, n, 1).setValues(code);
  sh.getRange(2, C.QR, n, 1).setFormulas(qr);
  SpreadsheetApp.flush();
  makePdfs_('qr');                                   // creates the PDFs and fills the QR PDF column
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// Scanner calls: ?action=scan&token=..&key=..&by=..   ?action=list&key=..   ?action=stats&key=..
function doGet(e) {
  const p = e.parameter || {};
  if (!p.action) return ContentService.createTextOutput('Check-in is running.');
  if (p.key !== staffKey_()) return json_({ result: 'denied' });
  if (p.action === 'scan') return json_(scan_(p.token, p.by));
  if (p.action === 'list') return json_(list_());
  if (p.action === 'stats') return json_(stats_());
  return json_({ result: 'unknown' });
}

function scan_(raw, by) {
  const token = String(raw || '').split('|').pop().trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    if (!token || last < 2) return { result: 'invalid' };
    const hit = sh.getRange(2, C.CODE, last - 1, 1).createTextFinder(token).matchEntireCell(true).findNext();
    if (!hit) return { result: 'invalid' };
    const row = hit.getRow();
    const r = sh.getRange(row, 1, 1, HEAD.length).getValues()[0];
    const who = { name: r[C.NAME - 1], designation: r[C.DESIG - 1], organization: r[C.ORG - 1] };
    if (r[C.IN - 1]) return Object.assign({ result: 'already', checkedInAt: String(r[C.TIME - 1]) }, who);
    const t = nowNepal_();
    sh.getRange(row, C.IN, 1, 3).setValues([['YES', t, by || '']]);
    try {
      logCheckin_(r, t, by);
    } catch (err) {}      // a log problem must never block a check-in
    return Object.assign({ result: 'ok' }, who);
  } finally {
    lock.releaseLock();
  }
}

// Compact attendee list for the scanner page: [code, name, designation, organization, checkedInTime]
// checkedInTime is '' when the person has not arrived yet. Phone and email are NOT sent.
function list_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return { result: 'list', rows: [] };
  const v = sh.getRange(2, 1, n, C.CODE).getValues();      // columns 1..Ticket code
  const rows = [];
  v.forEach(r => {
    const code = String(r[C.CODE - 1] || '').trim();
    if (!code) return;
    const t = r[C.IN - 1] ? (String(r[C.TIME - 1]) || 'earlier') : '';
    rows.push([code, r[C.NAME - 1], r[C.DESIG - 1], r[C.ORG - 1], t]);
  });
  return { result: 'list', rows: rows };
}

// Reads only the Checked in ... Ticket code columns instead of all 14.
function stats_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return { total: 0, checkedIn: 0, remaining: 0 };
  const v = sh.getRange(2, C.IN, n, C.CODE - C.IN + 1).getValues();
  let total = 0, checkedIn = 0;
  v.forEach(r => {
    if (r[C.CODE - C.IN]) {
      total++;
      if (r[0]) checkedIn++;
    }
  });
  return { total: total, checkedIn: checkedIn, remaining: total - checkedIn };
}

// ---------- Live check-in list: only people who have arrived, newest on top ----------
function liveSheet_() {
  const ss = SpreadsheetApp.getActive();
  let lg = ss.getSheetByName('Live Check-in');
  if (!lg) {
    lg = ss.insertSheet('Live Check-in');
    lg.getRange(1, 1, 1, 7)
      .setValues([['No.', 'Time (Nepal)', 'Name', 'Organization', 'Designation', 'Phone', 'Scanned by']])
      .setFontWeight('bold')
      .setBackground('#e8ebf2');
    lg.setFrozenRows(1);
    lg.setColumnWidth(2, 190);
    lg.setColumnWidth(3, 200);
    lg.setColumnWidth(4, 200);
  }
  return lg;
}

// Fewer sheet calls: format and values are written on one range (was 8 calls, now 4).
function logCheckin_(r, time, by) {
  const lg = liveSheet_();
  const no = lg.getLastRow();                       // header counts as 1, so this is the next number
  lg.insertRowBefore(2);
  lg.getRange(2, 1, 1, 7)
    .clearFormat()
    .setNumberFormats([['0', '@', '@', '@', '@', '@', '@']])   // text, so time and phone stay exact
    .setValues([[no, time, r[C.NAME - 1], r[C.ORG - 1], r[C.DESIG - 1], r[C.PHONE - 1], by || '']]);
}

// Typing YES by hand in the Checked in column also fills the time and adds the person to the live list.
function onEdit(e) {
  const rg = e.range, sh = rg.getSheet();
  if (sh.getName() !== 'Attendees' || rg.getRow() < 2 || rg.getColumn() > C.IN || rg.getLastColumn() < C.IN) return;
  for (let row = rg.getRow(); row <= rg.getLastRow(); row++) {
    const cell = sh.getRange(row, C.IN);
    if (String(cell.getValue()).trim().toUpperCase() !== 'YES' || sh.getRange(row, C.TIME).getValue()) continue;
    const r = sh.getRange(row, 1, 1, HEAD.length).getValues()[0];
    if (!r[C.NAME - 1]) continue;
    const t = nowNepal_();
    cell.setValue('YES');
    sh.getRange(row, C.TIME, 1, 2).setValues([[t, 'Manual']]);
    logCheckin_(r, t, 'Manual');
  }
}

// ---------- PDFs: invoices and QR tickets (one per attendee, saved in Google Drive) ----------
function esc_(t) {
  return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function pad_(n) {
  return ('000' + n).slice(-4);
}

function safe_(t) {
  return String(t).replace(/[\\\/:*?"<>|]/g, '').trim();
}

function logoUri_(s) {
  try {
    if (!s.logoDriveFileId) return '';

    // Automatically extract just the ID if the full URL was pasted
    let id = String(s.logoDriveFileId).trim();
    if (id.indexOf('/d/') > -1) {
      id = id.split('/d/')[1].split('/')[0];
    } else if (id.indexOf('id=') > -1) {
      id = id.split('id=')[1].split('&')[0];
    }

    const b = DriveApp.getFileById(id).getBlob();
    return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) {
    SpreadsheetApp.getUi().alert('Logo Error: Could not read the image.\n\nPlease ensure:\n1. The Google Drive link is correct in Settings.\n2. The image sharing is set to "Anyone with the link".');
    return '';
  }
}

// Same QR as the sheet link, embedded in the PDF. Tries twice in case the QR service is busy.
function qrUri_(text) {
  const url = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' + encodeURIComponent(text);
  for (let i = 0; i < 2; i++) {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) return 'data:image/png;base64,' + Utilities.base64Encode(res.getContent());
    Utilities.sleep(500);
  }
  return '';
}

function folder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function invoiceHtml_(r, s, logo, qr) {
  const no = 'INV-' + pad_(r[C.SERIAL - 1]);
  const date = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  const line = (a, b) => b
    ? '<tr><td style="padding:4px 14px 4px 0;color:#666;width:100px;">' + a + '</td><td style="font-weight:500;">' + esc_(b) + '</td></tr>'
    : '';

  // This puts the logo exactly in the top left area
  const logoHtml = logo ? '<img src="' + logo + '" height="80" style="display:block;margin-bottom:12px;">' : '';

  return '<div style="font-family:Arial,sans-serif;font-size:13px;color:#222;max-width:700px;margin:0 auto;">'

    // --- HEADER WITH LOGO ON LEFT ---
    + '<table width="100%"><tr>'
    + '<td valign="top" width="60%">'
    + logoHtml
    + '<b style="font-size:16px;">' + esc_(s.orgName) + '</b><br>'
    + '<span style="color:#555;">' + esc_(s.address) + '</span><br>'
    + '<span style="color:#555;">' + esc_(s.contact) + '</span>'
    + (s.panVat ? '<br><span style="color:#555;">PAN/VAT: ' + esc_(s.panVat) + '</span>' : '')
    + '</td>'
    + '<td align="right" valign="top" width="40%">'
    + '<b style="font-size:28px;color:#000;">INVOICE</b><br>'
    + '<span style="font-size:14px;color:#555;">' + no + '</span><br>'
    + '<span style="font-size:14px;color:#555;">Date: ' + date + '</span>'
    + '</td>'
    + '</tr></table>'

    // --- DETAILS SECTION ---
    + '<hr style="border:0;border-top:1px solid #ccc;margin:20px 0;">'
    + '<table width="100%" style="margin-top:14px;"><tr><td valign="top" width="55%">'

    // Participants details
    + '<h3 style="margin:0 0 10px 0;font-size:15px;color:#000;">Participants details</h3>'
    + '<table style="width:100%;">'
    + line('Name', r[C.NAME - 1])
    + line('Organization', r[C.ORG - 1])
    + line('Designation', r[C.DESIG - 1])
    + line('Phone', r[C.PHONE - 1])
    + line('Email', r[C.EMAIL - 1])
    + line('Check-in Time', r[C.TIME - 1])
    + line('Scanned By', r[C.BY - 1])
    + '</table>'

    // Event details
    + '<h3 style="margin:20px 0 10px 0;font-size:15px;color:#000;">Event details</h3>'
    + '<table style="width:100%;">'
    + line('Event', s.eventName)
    + line('Date', s.eventDate)
    + line('Venue', s.venue)
    + '</table>'

    // QR Code section on the right
    + '</td><td align="right" valign="top" width="45%">'
    + '<img src="' + qr + '" width="150" height="150" style="border:1px solid #eee;border-radius:8px;padding:4px;"><br>'
    + '<span style="font-size:11px;color:#777;">Ticket code: ' + esc_(r[C.CODE - 1]) + '</span>'
    + '</td></tr></table>'

    // --- FOOTER ---
    + '<hr style="border:0;border-top:1px solid #ccc;margin:25px 0 15px 0;">'
    + '<p style="margin-top:10px;color:#666;font-size:12px;text-align:center;">' + esc_(s.invoiceNote) + '</p>'
    + '</div>';
}

function ticketHtml_(r, s, logo, qr) {
  // The logo parameter is ignored here; it is only used for invoices.
  const line = (t, css) => t ? '<div style="' + css + '">' + esc_(t) + '</div>' : '';
  return '<div style="font-family:Arial,sans-serif;text-align:center;color:#222;padding:20px">'
    + line(s.orgName, 'font-size:14px;color:#555;margin-top:6px')
    + line(s.eventName, 'font-size:28px;font-weight:bold;margin:8px 0')
    + line([s.eventDate, s.venue].filter(String).join('  |  '), 'font-size:13px;color:#555')
    + '<div style="margin:22px 0 10px"><img src="' + qr + '" width="250" height="250"></div>'
    + line(r[C.NAME - 1], 'font-size:30px;font-weight:bold')
    + line(r[C.DESIG - 1], 'font-size:17px;color:#333;margin-top:4px')
    + line(r[C.ORG - 1], 'font-size:16px;color:#444;margin-top:2px')
    + line([r[C.PHONE - 1], r[C.EMAIL - 1]].filter(String).join('  |  '), 'font-size:14px;color:#555;margin-top:8px')
    + line(r[C.OTHER - 1], 'font-size:14px;color:#555;margin-top:2px')
    + line('Ticket code: ' + r[C.CODE - 1], 'font-size:12px;color:#777;margin-top:14px')
    + line('Show this QR code at the entrance', 'font-size:12px;color:#777;margin-top:4px')
    + '</div>';
}

function createInvoices() {
  makePdfs_('invoice');
}

// Makes a PDF for every row that has a name and a ticket code but no PDF link yet.
function makePdfs_(kind) {
  const sh = sheet_(), s = settings_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return SpreadsheetApp.getUi().alert('Add attendees first.');
  const isInv = kind === 'invoice';
  const col = isInv ? C.INV : C.QRPDF;
  const folderName = (isInv ? 'Invoices - ' : 'QR PDFs - ') + SpreadsheetApp.getActive().getName();
  const folder = folder_(folderName);
  const logo = logoUri_(s);
  const v = sh.getRange(2, 1, n, HEAD.length).getValues();
  const start = Date.now();
  let made = 0, noCode = 0, failed = 0, more = false, notCheckedIn = 0;

  for (let i = 0; i < v.length; i++) {
    const r = v[i];

    // Skip if no name, or already has a PDF
    if (!r[C.NAME - 1] || r[col - 1]) continue;

    // If generating an invoice, check if they are checked in (YES). If not, skip.
    if (isInv && String(r[C.IN - 1]).trim().toUpperCase() !== 'YES') {
      notCheckedIn++;
      continue;
    }

    if (!r[C.CODE - 1]) {
      noCode++;
      continue;
    }
    if (Date.now() - start > 270000) {      // stay under Google's 6 minute limit
      more = true;
      break;
    }
    const qr = qrUri_(qrData_(r));
    if (!qr) {
      failed++;
      continue;
    }

    // Pass logo only to invoice generation, leave ticket generation clean
    const html = isInv ? invoiceHtml_(r, s, logo, qr) : ticketHtml_(r, s, '', qr);
    const name = (isInv ? 'Invoice' : 'QR') + ' - ' + pad_(r[C.SERIAL - 1]) + ' - ' + safe_(r[C.NAME - 1]) + '.pdf';
    const f = folder.createFile(HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(name));
    sh.getRange(i + 2, col).setFormula('=HYPERLINK("' + f.getUrl() + '","Open PDF")');
    made++;
  }

  SpreadsheetApp.getUi().alert(made + ' PDFs created in Drive folder "' + folderName + '".'
    + (noCode ? '\n' + noCode + ' skipped: run "Generate QR codes + PDFs" first.' : '')
    + (notCheckedIn ? '\n' + notCheckedIn + ' skipped: Not checked in yet (Invoices only).' : '')
    + (failed ? '\n' + failed + ' failed (QR service busy). Run again.' : '')
    + (more ? '\n' + (isInv ? 'Run again for the rest.' : 'Click "Generate QR codes + PDFs" again for the rest.') : ''));
}

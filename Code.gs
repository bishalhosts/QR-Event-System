// Event check-in on Google Sheets. Reuse for any event: File > Make a copy of this Sheet.
const TZ = 'Asia/Kathmandu';   // Nepal time
const HEAD = ['Serial No', 'Name', 'Organization', 'Designation', 'Phone', 'Email', 'Other',
  'Ticket code', 'QR code (download link)', 'Checked in', 'Check-in time (Nepal)', 'Scanned by', 'Amount', 'Invoice PDF'];
const C = { SERIAL: 1, NAME: 2, ORG: 3, DESIG: 4, PHONE: 5, EMAIL: 6, OTHER: 7, CODE: 8, QR: 9, IN: 10, TIME: 11, BY: 12, AMT: 13, INV: 14 };

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Check-in')
    .addItem('1. Set up sheets', 'setup')
    .addItem('2. Generate QR codes', 'generateQR')
    .addItem('3. Create invoice PDFs', 'createInvoices')
    .addToUi();
}

function sheet_() { return SpreadsheetApp.getActive().getSheetByName('Attendees'); }
function nowNepal_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd hh:mm:ss a'); }

function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Attendees') || ss.insertSheet('Attendees');
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold').setBackground('#e8ebf2');
  sh.setFrozenRows(1);
  const rows = sh.getMaxRows() - 1;
  sh.getRange(2, C.TIME, rows, 1).setNumberFormat('@');          // keep time as exact text
  sh.setColumnWidth(C.QR, 170); sh.setColumnWidth(C.TIME, 190);
  // Green when Checked in = YES (strong green on the YES cell, light green on the row)
  const yes = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$J2="YES"')
    .setBackground('#1f9d55').setFontColor('#ffffff').setBold(true).setRanges([sh.getRange(2, C.IN, rows, 1)]).build();
  const row = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$J2="YES"')
    .setBackground('#d5efdf').setRanges([sh.getRange(2, 1, rows, HEAD.length)]).build();
  sh.setConditionalFormatRules([yes, row]);

  const st = ss.getSheetByName('Settings') || ss.insertSheet('Settings');
  if (st.getLastRow() === 0) {
    st.getRange(1, 1, 13, 3).setValues([
      ['Setting', 'Value', 'Note'],
      ['staffKey', Utilities.getUuid().slice(0, 8), 'Password the scanner asks for'],
      ['orgName', 'My Organization', 'Shown on invoices'],
      ['eventName', 'My Event', 'Shown on invoices'],
      ['eventDate', '', 'Shown on invoices'],
      ['venue', '', 'Shown on invoices'],
      ['address', '', 'Organization address'],
      ['contact', '', 'Phone / email on invoice'],
      ['panVat', '', 'PAN / VAT number (optional)'],
      ['logoDriveFileId', '', 'Upload logo to Google Drive, paste its file ID (between /d/ and /view in the link)'],
      ['currency', 'NPR', ''],
      ['defaultAmount', 0, 'Used when Amount column is empty'],
      ['invoiceNote', 'Thank you for registering.', 'Footer text']
    ]);
    st.getRange(1, 1, 1, 3).setFontWeight('bold');
    st.setColumnWidth(2, 260); st.setColumnWidth(3, 420);
  }
  SpreadsheetApp.getUi().alert('Ready. Add attendees from row 2 (Name is required), edit the Settings sheet, then run "Generate QR codes".');
}

function settings_() {
  const o = {};
  SpreadsheetApp.getActive().getSheetByName('Settings').getDataRange().getValues().slice(1).forEach(r => o[r[0]] = r[1]);
  return o;
}

// Serial numbers, ticket codes and QR download links for every row that has a name.
function generateQR() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return SpreadsheetApp.getUi().alert('Add attendees first (Name column).');
  const v = sh.getRange(2, 1, n, HEAD.length).getValues();
  const serial = [], code = [], qr = [];
  v.forEach((r, i) => {
    const row = i + 2;
    if (!r[C.NAME - 1]) { serial.push(['']); code.push(['']); qr.push(['']); return; }
    serial.push([r[C.SERIAL - 1] || i + 1]);
    code.push([r[C.CODE - 1] || Utilities.getUuid().replace(/-/g, '').slice(0, 12)]);
    // QR text = Name | Designation | Code  (a normal phone camera shows name and designation)
    qr.push(['=HYPERLINK("https://api.qrserver.com/v1/create-qr-code/?size=400x400&download=1&data="&ENCODEURL(B' + row + '&" | "&D' + row + '&" | "&H' + row + '),"Download QR")']);
  });
  sh.getRange(2, C.SERIAL, n, 1).setValues(serial);
  sh.getRange(2, C.CODE, n, 1).setValues(code);
  sh.getRange(2, C.QR, n, 1).setFormulas(qr);
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Scanner calls: ?action=scan&token=..&key=..&by=..   and   ?action=stats&key=..
function doGet(e) {
  const p = e.parameter || {};
  if (!p.action) return ContentService.createTextOutput('Check-in is running.');
  if (p.key !== String(settings_().staffKey)) return json_({ result: 'denied' });
  if (p.action === 'scan') return json_(scan_(p.token, p.by));
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
    sh.getRange(row, C.IN, 1, 3).setValues([['YES', nowNepal_(), by || '']]);
    return Object.assign({ result: 'ok' }, who);
  } finally { lock.releaseLock(); }
}

function stats_() {
  const sh = sheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return { total: 0, checkedIn: 0, remaining: 0 };
  const v = sh.getRange(2, C.CODE, n, 3).getValues().filter(r => r[0]);
  const checkedIn = v.filter(r => r[2]).length;
  return { total: v.length, checkedIn: checkedIn, remaining: v.length - checkedIn };
}

// ---------- Invoices (one PDF per attendee, saved in Google Drive) ----------
function esc_(t) { return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

function invoiceHtml_(r, s, logo) {
  const amt = Number(r[C.AMT - 1] || s.defaultAmount || 0);
  const no = 'INV-' + ('000' + r[C.SERIAL - 1]).slice(-4);
  const date = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const line = (a, b) => b ? '<tr><td style="padding:4px 12px 4px 0;color:#666">' + a + '</td><td>' + esc_(b) + '</td></tr>' : '';
  return '<div style="font-family:Arial,sans-serif;font-size:13px;color:#222;padding:10px">'
    + '<table width="100%"><tr><td>' + (logo ? '<img src="' + logo + '" height="60">' : '') + '</td>'
    + '<td align="right"><b style="font-size:24px">INVOICE</b><br>' + no + '<br>Date: ' + date + '</td></tr></table><hr>'
    + '<b style="font-size:15px">' + esc_(s.orgName) + '</b><br>' + esc_(s.address) + '<br>' + esc_(s.contact)
    + (s.panVat ? '<br>PAN/VAT: ' + esc_(s.panVat) : '')
    + '<h3>Bill to</h3><table>' + line('Name', r[C.NAME - 1]) + line('Organization', r[C.ORG - 1]) + line('Designation', r[C.DESIG - 1])
    + line('Phone', r[C.PHONE - 1]) + line('Email', r[C.EMAIL - 1]) + '</table>'
    + '<h3>Event</h3><table>' + line('Event', s.eventName) + line('Date', s.eventDate) + line('Venue', s.venue) + '</table><br>'
    + '<table width="100%" style="border-collapse:collapse"><tr style="background:#eee"><td style="padding:8px"><b>Description</b></td><td align="right" style="padding:8px"><b>Amount</b></td></tr>'
    + '<tr><td style="padding:8px;border-bottom:1px solid #ccc">' + esc_(s.eventName) + ' registration</td><td align="right" style="padding:8px;border-bottom:1px solid #ccc">' + esc_(s.currency) + ' ' + amt.toFixed(2) + '</td></tr>'
    + '<tr><td align="right" style="padding:8px"><b>Total</b></td><td align="right" style="padding:8px"><b>' + esc_(s.currency) + ' ' + amt.toFixed(2) + '</b></td></tr></table>'
    + '<p style="margin-top:30px;color:#555">' + esc_(s.invoiceNote) + '</p></div>';
}

function createInvoices() {
  const sh = sheet_(), s = settings_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  let logo = '';
  try {
    if (s.logoDriveFileId) { const b = DriveApp.getFileById(String(s.logoDriveFileId)).getBlob(); logo = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); }
  } catch (e) { SpreadsheetApp.getActive().toast('Logo could not be read. Check the logo file ID in Settings.'); }
  const folderName = 'Invoices - ' + SpreadsheetApp.getActive().getName();
  const it = DriveApp.getFoldersByName(folderName);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
  const v = sh.getRange(2, 1, n, HEAD.length).getValues();
  const start = Date.now();
  let made = 0, more = false;
  for (let i = 0; i < v.length; i++) {
    const r = v[i];
    if (!r[C.NAME - 1] || !r[C.SERIAL - 1] || r[C.INV - 1]) continue;
    if (Date.now() - start > 270000) { more = true; break; }       // stay under Google's 6 minute limit
    const pdf = HtmlService.createHtmlOutput(invoiceHtml_(r, s, logo)).getBlob().getAs('application/pdf')
      .setName('Invoice-' + r[C.SERIAL - 1] + '-' + r[C.NAME - 1] + '.pdf');
    const f = folder.createFile(pdf);
    sh.getRange(i + 2, C.INV).setFormula('=HYPERLINK("' + f.getUrl() + '","Open PDF")');
    made++;
  }
  SpreadsheetApp.getUi().alert(made + ' invoices created in Drive folder "' + folderName + '".' + (more ? ' Run again for the rest.' : ''));
}

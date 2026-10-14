/**
 * Nimbus SIM journey: Google Sheets backend (Google Apps Script)
 *
 * Paste into Extensions > Apps Script of your Google Sheet, run setup() once,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */

const TABS = {
  applicants: {
    name: 'Applicants',
    headers: ['Reference ID', 'Submitted at', 'Started at', 'Mobile', 'Pincode', 'OTP verified',
      'Email', 'First name', 'Last name', 'Aadhaar (masked)', 'Aadhaar hash', 'Age confirmed',
      'Outcome', 'Notice version', 'User agent'],
    textCols: [1, 4, 5, 10, 11]           // keep leading zeros / long digits as text
  },
  consents: {
    name: 'Consents',
    headers: ['Reference ID', 'Mobile', 'Purpose', 'Item', 'Status', 'Captured at', 'Last updated', 'Notice version'],
    textCols: [1, 2]
  },
  events: {
    name: 'Events',
    headers: ['Reference ID', 'Step', 'Action', 'Occurred at'],
    textCols: [1]
  },
  audit: {
    name: 'Audit',
    headers: ['Changed at', 'Changed by', 'Reference ID', 'Purpose', 'Item', 'Old status', 'New status', 'Consents row'],
    textCols: [3]
  }
};

const STATUS_COL = 5;                       // "Status" column in Consents
const STATUSES = ['granted', 'declined', 'withdrawn'];
const ITEMS = {
  privacy_notice: ['service_processing'],
  marketing: ['calls', 'whatsapp', 'sms', 'email'],
  cross_service: ['data_linking'],
  interest: ['shopping_lifestyle', 'entertainment', 'sports_experiences', 'financial_services', 'health_wellness', 'education']
};
const OUTCOMES = ['completed', 'marketing_declined', 'cross_service_declined'];

/* ---------- One-time setup: run this from the editor ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Kolkata');
  Object.values(TABS).forEach(t => ensureTab_(ss, t));

  const consents = ss.getSheetByName(TABS.consents.name);
  consents.getRange(2, STATUS_COL, consents.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build());

  const audit = ss.getSheetByName(TABS.audit.name);
  if (audit.getProtections(SpreadsheetApp.ProtectionType.SHEET).length === 0) {
    audit.protect().setDescription('Audit log: written by script').setWarningOnly(true);
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('AADHAAR_SALT')) {
    props.setProperty('AADHAAR_SALT', Utilities.getUuid() + Utilities.getUuid());
  }

  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);
}

function ensureTab_(ss, t) {
  let sh = ss.getSheetByName(t.name);
  if (!sh) sh = ss.insertSheet(t.name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, t.headers.length).setValues([t.headers]).setFontWeight('bold').setBackground('#E6E8F7');
    sh.setFrozenRows(1);
    t.textCols.forEach(c => sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'));
    sh.autoResizeColumns(1, t.headers.length);
  }
  return sh;
}

/* ---------- Web app endpoints ---------- */
function doGet() {
  return json_({ ok: true, message: 'Nimbus SIM sheet endpoint is running. Send records with POST.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    let p;
    try {
      p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    } catch (err) {
      return json_({ ok: false, error: 'The request body must be valid JSON.' });
    }

    const errors = validate_(p);
    if (errors.length) return json_({ ok: false, error: 'Some fields are invalid.', details: errors });

    const salt = PropertiesService.getScriptProperties().getProperty('AADHAAR_SALT');
    if (!salt) return json_({ ok: false, error: 'The sheet is not set up yet. Run setup() in Apps Script.' });

    const a = p.applicant;
    const digits = String(a.aadhaar).replace(/\s/g, '');
    const masked = 'XXXX XXXX ' + digits.slice(-4);
    const hash = hmacHex_(digits, salt);
    const ref = referenceId_();
    const submittedAt = new Date();
    const notice = str_(p.noticeVersion);

    lock.waitLock(20000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    append_(ensureTab_(ss, TABS.applicants), [[
      ref, submittedAt, toDate_(p.startedAt), a.mobile, a.pincode || '', true,
      str_(a.email).toLowerCase(), str_(a.firstName), str_(a.lastName), masked, hash, true,
      p.outcome, notice, str_(p.client && p.client.userAgent).slice(0, 300)
    ]]);

    append_(ensureTab_(ss, TABS.consents), p.consents.map(c => [
      ref, a.mobile, c.purpose, c.item, c.status, toDate_(c.capturedAt), submittedAt, notice
    ]));

    const events = (p.events || [])
      .filter(ev => ev && str_(ev.step) && str_(ev.action) && isDate_(ev.at))
      .slice(0, 200)
      .map(ev => [ref, str_(ev.step).slice(0, 60), str_(ev.action).slice(0, 60), toDate_(ev.at)]);
    if (events.length) append_(ensureTab_(ss, TABS.events), events);

    SpreadsheetApp.flush();
    return json_({ ok: true, referenceId: ref });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: "The application couldn't be saved. Try again in a moment." });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Audit trail for manual consent edits ---------- */
// Runs automatically when someone edits the sheet by hand.
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (sh.getName() !== TABS.consents.name) return;
  const r = e.range;
  if (r.getRow() < 2 || r.getColumn() > STATUS_COL || r.getLastColumn() < STATUS_COL) return;

  const ss = sh.getParent();
  const audit = ensureTab_(ss, TABS.audit);
  let who = 'unknown';
  try { who = (e.user && e.user.getEmail()) || Session.getActiveUser().getEmail() || 'unknown'; } catch (err) {}

  const now = new Date();
  const rows = [];
  for (let i = 0; i < r.getNumRows(); i++) {
    const row = r.getRow() + i;
    const [ref, , purpose, item, status] = sh.getRange(row, 1, 1, STATUS_COL).getValues()[0];
    const single = r.getNumRows() === 1 && r.getNumColumns() === 1;
    rows.push([now, who, ref, purpose, item, single ? (e.oldValue || '') : '(multi-cell edit)', status, row]);
    sh.getRange(row, 7).setValue(now);      // "Last updated"
  }
  append_(audit, rows);
}

/* ---------- Helpers ---------- */
function validate_(p) {
  const e = [];
  const a = (p && p.applicant) || {};
  if (!/^[6-9]\d{9}$/.test(str_(a.mobile))) e.push('applicant.mobile must be a 10-digit Indian mobile number');
  if (a.pincode && !/^[1-9]\d{5}$/.test(str_(a.pincode))) e.push('applicant.pincode must be 6 digits');
  if (a.otpVerified !== true) e.push('applicant.otpVerified must be true');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(str_(a.email))) e.push('applicant.email is invalid');
  if (!str_(a.firstName) || str_(a.firstName).length > 50) e.push('applicant.firstName is required');
  if (!str_(a.lastName) || str_(a.lastName).length > 50) e.push('applicant.lastName is required');
  if (!/^\d{12}$/.test(str_(a.aadhaar).replace(/\s/g, ''))) e.push('applicant.aadhaar must be 12 digits');
  if (a.ageConfirmed !== true) e.push('applicant.ageConfirmed must be true');
  if (OUTCOMES.indexOf(p && p.outcome) < 0) e.push('outcome is invalid');
  if (!str_(p && p.noticeVersion)) e.push('noticeVersion is required');
  if (!Array.isArray(p && p.consents) || p.consents.length === 0 || p.consents.length > 50) {
    e.push('consents must be a non-empty list');
  } else {
    p.consents.forEach((c, i) => {
      if (!c || !ITEMS[c.purpose] || ITEMS[c.purpose].indexOf(c.item) < 0) e.push('consents[' + i + '] has an unknown purpose or item');
      if (!c || ['granted', 'declined'].indexOf(c.status) < 0) e.push('consents[' + i + '].status must be granted or declined');
      if (!c || !isDate_(c.capturedAt)) e.push('consents[' + i + '].capturedAt must be an ISO date');
    });
  }
  return e;
}

function append_(sh, rows) {
  if (!rows.length) return;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}
function referenceId_() {
  const day = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd');
  return 'NMB-' + day + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
}
function hmacHex_(value, key) {
  return Utilities.computeHmacSha256Signature(value, key)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function str_(v) { return typeof v === 'string' ? v.trim() : ''; }
function isDate_(v) { return typeof v === 'string' && !isNaN(Date.parse(v)); }
function toDate_(v) { return isDate_(v) ? new Date(v) : ''; }

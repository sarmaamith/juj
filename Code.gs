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
  securiti: {
    name: 'Securiti sync',
    headers: ['Reference ID', 'Attempted at', 'Result', 'HTTP status', 'Response', 'Request body'],
    textCols: [1]
  },
  inbound: {
    name: 'Securiti inbound',
    headers: ['Received at', 'UUID', 'Trigger', 'Consent source', 'Result', 'Details'],
    textCols: [2]
  },
  audit: {
    name: 'Audit',
    headers: ['Changed at', 'Changed by', 'Reference ID', 'Purpose', 'Item', 'Old status', 'New status', 'Consents row'],
    textCols: [3]
  }
};

/* ---------- Securiti consent upload ----------
 * Keys are NOT stored here. Add them in Apps Script > Project Settings > Script properties:
 *   SECURITI_API_KEY, SECURITI_API_SECRET, SECURITI_TIDENT
 *
 * Each website consent is mapped to a Securiti processing purpose and consent purpose.
 * The names must match the purposes configured in your Securiti tenant exactly.
 */
const SECURITI = {
  url: 'https://app.securiti.ai/privaci/v1/consentapi/consent_upload',
  consentSource: 'Nimbus SIM Application',
  purposeMap: {
    'privacy_notice.service_processing': { processing: 'Service Processing',      item: 'Service Activation' },
    'marketing.calls':                   { processing: 'Marketing Communication', item: 'Call' },
    'marketing.whatsapp':                { processing: 'Marketing Communication', item: 'WhatsApp' },
    'marketing.sms':                     { processing: 'Marketing Communication', item: 'SMS' },
    'marketing.email':                   { processing: 'Marketing Communication', item: 'Email' },
    'cross_service.data_linking':        { processing: 'Connected Experiences',   item: 'Use Information Across Services' },
    'interest.shopping_lifestyle':       { processing: 'Group Offers',            item: 'Shopping and Lifestyle' },
    'interest.entertainment':            { processing: 'Group Offers',            item: 'Entertainment' },
    'interest.sports_experiences':       { processing: 'Group Offers',            item: 'Sports and Experiences' },
    'interest.financial_services':       { processing: 'Group Offers',            item: 'Financial Services' },
    'interest.health_wellness':          { processing: 'Group Offers',            item: 'Health and Wellness' },
    'interest.education':                { processing: 'Group Offers',            item: 'Education' }
  }
};

/**
 * Builds the consent_upload body in Securiti's schema. Only the identifiers and the
 * consent choices are sent: no name, Aadhaar, pincode or other application details.
 * Every consent the customer saw is sent as true or false, including declined ones.
 */
function buildSecuritiBody_(applicant, consents) {
  const email = str_(applicant.email).toLowerCase();
  const phone = '+91' + applicant.mobile;

  const groups = {};          // processing purpose -> consented items, in journey order
  consents.forEach(c => {
    const m = SECURITI.purposeMap[c.purpose + '.' + c.item];
    if (!m) return;
    (groups[m.processing] = groups[m.processing] || []).push({
      consent_purpose_name: m.item,
      consent_granted: c.status === 'granted'
    });
  });

  return {
    UUID: email,
    uuids: { email: email, phone: phone },
    activity_timestamp: Math.floor(Date.now() / 1000),
    consent_source: SECURITI.consentSource,
    consent_info: Object.keys(groups).map(name => ({
      processing_purpose_name: name,
      consented_items: groups[name]
    }))
  };
}

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
  if (!props.getProperty('INBOUND_KEY')) {
    props.setProperty('INBOUND_KEY', Utilities.getUuid().replace(/-/g, ''));
  }
  console.log('Securiti webhook URL: <your web app URL>?key=' + props.getProperty('INBOUND_KEY'));

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

    // "Manage preferences" on the website: find emails for a mobile number
    if (p && p.action === 'lookup_emails') return json_(lookupEmails_(p.mobile));

    // Consent changes pushed by Securiti (preference center) take a separate path
    if (isSecuritiEvent_(p)) {
      lock.waitLock(20000);
      return json_(applySecuritiEvents_(e, p));
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
    lock.releaseLock();   // don't hold the sheet while waiting on Securiti

    const securiti = syncToSecuriti_(ref, a, p.consents, notice);
    return json_({ ok: true, referenceId: ref, securiti: securiti.result });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: "The application couldn't be saved. Try again in a moment." });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Securiti sync ---------- */
function syncToSecuriti_(ref, applicant, consents, noticeVersion) {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECURITI_API_KEY') || !props.getProperty('SECURITI_API_SECRET') || !props.getProperty('SECURITI_TIDENT')) {
    return logSync_(ref, 'not_configured', '', 'Securiti keys are not set in Script properties.', '');
  }
  const unmapped = consents.filter(c => !SECURITI.purposeMap[c.purpose + '.' + c.item]);
  if (unmapped.length) {
    return logSync_(ref, 'not_configured', '', 'No Securiti purpose mapped for: ' +
      unmapped.map(c => c.purpose + '.' + c.item).join(', '), '');
  }
  return sendToSecuriti_(ref, JSON.stringify(buildSecuritiBody_(applicant, consents)));
}

function sendToSecuriti_(ref, body) {
  const props = PropertiesService.getScriptProperties();
  try {
    const res = UrlFetchApp.fetch(SECURITI.url, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'X-API-Key': props.getProperty('SECURITI_API_KEY'),
        'X-API-Secret': props.getProperty('SECURITI_API_SECRET'),
        'X-TIDENT': props.getProperty('SECURITI_TIDENT')
      },
      payload: body,
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const ok = code >= 200 && code < 300;
    return logSync_(ref, ok ? 'sent' : 'failed', code, res.getContentText().slice(0, 1000), body);
  } catch (err) {
    return logSync_(ref, 'failed', '', String(err).slice(0, 1000), body);
  }
}

function logSync_(ref, result, code, response, body) {
  try {
    const sh = ensureTab_(SpreadsheetApp.getActiveSpreadsheet(), TABS.securiti);
    append_(sh, [[ref, new Date(), result, code, response, body]]);
  } catch (err) { console.error(err); }
  return { result: result };
}

/**
 * Resends every submission whose latest Securiti attempt failed.
 * Run it from the editor, or add a time-driven trigger (e.g. hourly).
 */
function retryFailedSecuritiSyncs() {
  const sh = ensureTab_(SpreadsheetApp.getActiveSpreadsheet(), TABS.securiti);
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];
  const latest = {};
  rows.forEach(r => { latest[r[0]] = r; });   // last attempt per reference ID
  Object.values(latest)
    .filter(r => r[2] === 'failed' && r[5])
    .forEach(r => {
      // Resend the same choices with a fresh activity timestamp
      let body = r[5];
      try { const b = JSON.parse(body); b.activity_timestamp = Math.floor(Date.now() / 1000); body = JSON.stringify(b); } catch (err) {}
      sendToSecuriti_(r[0], body);
    });
}


/* ---------- Inbound: consent changes from the Securiti preference center ----------
 * Securiti posts events to:  <web app URL>?key=<INBOUND_KEY>
 * (Apps Script can't read request headers, so the key travels in the URL.)
 * Each event updates the matching rows in Consents and adds a row to Audit.
 * Nothing is sent back to Securiti, so there's no loop.
 */
// Customers who aren't in Applicants still get Consents rows, under the Reference ID
// "SECURITI:<email>". Set to false to only log them as unmatched.
const ADD_UNMATCHED_CUSTOMERS = true;

const INBOUND_STATUS = {
  GRANTED: 'granted', WITHDRAWN: 'withdrawn', REVOKED: 'withdrawn', OPTED_OUT: 'withdrawn',
  DECLINED: 'declined', DENIED: 'declined'
};

function isSecuritiEvent_(p) {
  const list = Array.isArray(p) ? p : [p];
  return list.length > 0 && list.every(ev => ev && ev.consented_items && (ev.UUID || ev.UUIDs || ev.user_uuid));
}

function applySecuritiEvents_(e, p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inbound = ensureTab_(ss, TABS.inbound);
  const key = PropertiesService.getScriptProperties().getProperty('INBOUND_KEY');
  if (!key || !e || !e.parameter || e.parameter.key !== key) {
    append_(inbound, [[new Date(), '', '', '', 'rejected', 'Missing or wrong key in the URL']]);
    return { ok: false, error: 'Unauthorized.' };
  }

  // Website consent key ("marketing.sms") for each Securiti name pair
  const reverse = {};
  Object.keys(SECURITI.purposeMap).forEach(k => {
    const m = SECURITI.purposeMap[k];
    reverse[(m.processing + '|' + m.item).toLowerCase()] = k;
  });

  const appSh = ensureTab_(ss, TABS.applicants);
  const apps = appSh.getLastRow() > 1 ? appSh.getRange(2, 1, appSh.getLastRow() - 1, 7).getValues() : [];
  const conSh = ensureTab_(ss, TABS.consents);
  const data = conSh.getLastRow() > 1 ? conSh.getRange(2, 1, conSh.getLastRow() - 1, 8).getValues() : [];
  const index = {};
  data.forEach((r, i) => { index[r[0] + '|' + r[2] + '|' + r[3]] = i; });
  const auditSh = ensureTab_(ss, TABS.audit);

  const events = Array.isArray(p) ? p : [p];
  const totals = { events: events.length, updated: 0, added: 0, unchanged: 0, stale: 0 };

  events.forEach(ev => {
    const id = identifiers_(ev);
    const refs = apps
      .filter(r => (id.email && String(r[6]).toLowerCase() === id.email) || (id.mobile && String(r[3]) === id.mobile))
      .map(r => ({ ref: String(r[0]), mobile: String(r[3]) }));
    const when = Number(ev.consent_timestamp) ? new Date(Number(ev.consent_timestamp) * 1000) : new Date();
    const source = 'Securiti ' + (ev.trigger_type || 'update');
    const notes = [];
    const audits = [];
    let changed = 0;

    let label = '';
    if (!refs.length) {
      if (ADD_UNMATCHED_CUSTOMERS && (id.email || id.mobile)) {
        refs.push({ ref: 'SECURITI:' + (id.email || id.mobile), mobile: id.mobile });
        label = ' (Securiti-only customer)';
      } else {
        notes.push('No applicant found for this email or mobile');
      }
    }

    Object.keys(ev.consented_items).forEach(k => {
      const it = ev.consented_items[k] || {};
      const ppName = (it.processing_purpose && it.processing_purpose.name) || '';
      const cpName = (it.consent_purpose && it.consent_purpose.name) || '';
      const mapped = reverse[(ppName + '|' + cpName).toLowerCase()];
      const status = INBOUND_STATUS[String(it.consent_status || '').toUpperCase()];
      if (!mapped) { notes.push('Not mapped: ' + ppName + ' / ' + cpName); return; }
      if (!status) { notes.push('Unknown status ' + it.consent_status + ' for ' + cpName); return; }
      const purpose = mapped.split('.')[0], item = mapped.split('.')[1];

      refs.forEach(a => {
        const i = index[a.ref + '|' + purpose + '|' + item];
        if (i === undefined) {
          const row = [a.ref, a.mobile, purpose, item, status, when, when, 'securiti'];
          append_(conSh, [row]);
          data.push(row);
          index[a.ref + '|' + purpose + '|' + item] = data.length - 1;
          audits.push([new Date(), source, a.ref, purpose, item, '(none)', status, data.length + 1]);
          totals.added++; changed++;
          return;
        }
        const row = data[i];
        if (row[4] === status) { totals.unchanged++; return; }
        if (row[6] instanceof Date && row[6] > when) { totals.stale++; return; }   // sheet has a newer change
        conSh.getRange(i + 2, STATUS_COL).setValue(status);
        conSh.getRange(i + 2, 7).setValue(when);
        audits.push([new Date(), source, a.ref, purpose, item, row[4], status, i + 2]);
        row[4] = status; row[6] = when;
        totals.updated++; changed++;
      });
    });

    if (audits.length) append_(auditSh, audits);
    append_(inbound, [[new Date(), id.email || id.mobile, ev.trigger_type || '', ev.consent_source || '',
      refs.length ? (changed ? changed + ' change(s)' : 'no changes') + label : 'unmatched',
      notes.join('; ').slice(0, 1000)]]);
  });

  SpreadsheetApp.flush();
  return Object.assign({ ok: true }, totals);
}

function identifiers_(ev) {
  const ids = {};
  if (ev.UUIDs && typeof ev.UUIDs === 'object') Object.assign(ids, ev.UUIDs);
  try { Object.assign(ids, JSON.parse(ev.user_uuid || '{}')); } catch (err) {}
  const uuid = String(ev.UUID || '');
  if (uuid.indexOf('@') > 0) ids.email = ids.email || uuid;
  else if (uuid) ids.phone = ids.phone || uuid;
  if (ev.email) ids.email = ids.email || ev.email;
  if (ev.phone_number) ids.phone = ids.phone || ev.phone_number;
  return {
    email: String(ids.email || '').trim().toLowerCase(),
    mobile: String(ids.phone || '').replace(/\D/g, '').slice(-10)
  };
}

/* ---------- Lookup for "Manage preferences" ---------- */
// The website is public, so emails are masked (r***@gmail.com) to stop anyone who types
// a stranger's number from seeing their email. Set to false to return full addresses.
const MASK_LOOKUP_EMAILS = true;

function lookupEmails_(mobile) {
  const m = String(mobile || '').replace(/\D/g, '');
  if (!/^[6-9]\d{9}$/.test(m)) return { ok: false, error: 'Enter a 10-digit mobile number starting with 6, 7, 8 or 9.' };
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TABS.applicants.name);
  const rows = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues() : [];
  const seen = {};
  rows.forEach(r => {
    const email = String(r[6] || '').trim().toLowerCase();
    if (String(r[3]) === m && email) seen[email] = true;
  });
  const emails = Object.keys(seen).map(e => MASK_LOOKUP_EMAILS ? maskEmail_(e) : e);
  return { ok: true, emails: emails };
}

function maskEmail_(email) {
  const at = email.indexOf('@');
  if (at < 1) return '***';
  return email.charAt(0) + '***' + email.slice(at);
}

/* ---------- Manual import: paste Securiti JSON into a dialog ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Securiti')
    .addItem('Import consent JSON…', 'showImportDialog')
    .addToUi();
}

function showImportDialog() {
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif">' +
    '<p style="margin:0 0 8px">Paste the JSON from Securiti (one event or a list of events), then click Import.</p>' +
    '<textarea id="t" style="width:100%;height:300px;font-family:monospace;font-size:12px"></textarea>' +
    '<p><button id="b" onclick="go()" style="padding:8px 18px">Import</button> <span id="m"></span></p>' +
    '<script>function go(){var b=document.getElementById("b"),m=document.getElementById("m");' +
    'b.disabled=true;m.textContent="Importing…";' +
    'google.script.run.withSuccessHandler(function(r){m.textContent=r;b.disabled=false;})' +
    '.withFailureHandler(function(e){m.textContent="Error: "+e.message;b.disabled=false;})' +
    '.importSecuritiJson(document.getElementById("t").value);}</script></div>'
  ).setWidth(640).setHeight(440);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Securiti consent JSON');
}

/** Parses pasted Securiti JSON and applies it exactly like a webhook call. */
function importSecuritiJson(text) {
  const p = parseSecuritiText_(text);
  if (!isSecuritiEvent_(p)) throw new Error('This doesn\'t look like Securiti consent JSON (needs UUID and consented_items).');
  const key = PropertiesService.getScriptProperties().getProperty('INBOUND_KEY');
  if (!key) throw new Error('Run setup() once first.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const r = applySecuritiEvents_({ parameter: { key: key } }, p);
    return 'Done: ' + r.events + ' event(s), ' + r.updated + ' updated, ' + r.added + ' added, ' +
      r.unchanged + ' unchanged, ' + r.stale + ' skipped as older. See the Securiti inbound tab.';
  } finally {
    lock.releaseLock();
  }
}

// Accepts clean JSON, or JSON copied from a browser viewer with stray "-" collapse markers
function parseSecuritiText_(text) {
  const raw = String(text || '').trim();
  try { return JSON.parse(raw); } catch (err) {}
  const cleaned = raw.replace(/(^|\s)-(?=\s|$)/g, '$1');
  try { return JSON.parse(cleaned); } catch (err) {
    throw new Error('Couldn\'t read the JSON. Copy it again using the viewer\'s "Raw" view if it has one.');
  }
}

/**
 * Try the inbound path from the editor: withdraws SMS consent for the most recent applicant,
 * the same way a Securiti preference-center update would. Check Consents, Audit and Securiti inbound after.
 */
function testSecuritiInbound() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const apps = ss.getSheetByName(TABS.applicants.name);
  if (!apps || apps.getLastRow() < 2) throw new Error('Submit an application on the website first.');
  const email = apps.getRange(apps.getLastRow(), 7).getValue();
  const sms = SECURITI.purposeMap['marketing.sms'];
  const event = [{
    UUID: email,
    UUIDs: { email: email },
    consent_source: 'Test',
    consent_timestamp: Math.floor(Date.now() / 1000),
    trigger_type: 'preference-center',
    consented_items: {
      test_1: { consent_status: 'WITHDRAWN', processing_purpose: { name: sms.processing }, consent_purpose: { name: sms.item } }
    }
  }];
  const key = PropertiesService.getScriptProperties().getProperty('INBOUND_KEY');
  console.log(JSON.stringify(applySecuritiEvents_({ parameter: { key: key } }, event)));
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

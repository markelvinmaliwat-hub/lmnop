/**
 * MMMDA Logistics Dashboard - Google Apps Script backend (Code.gs)
 *
 * Works with index.html (WEB_APP_URL / LOG_URL / SHEET_ID in the page).
 *
 * What it does
 *  - ?action=write&data={...}      saves an activity row, or a payment row when data.type === 'pay'
 *  - ?action=read&kind=activity    returns the activity log      -> {data:[...]}
 *  - ?action=read&kind=pay         returns latest payment per key -> {data:[...]}
 *  - ?action=read&kind=payhist     returns payment history        -> {data:[...]}
 *  - ?acct=create|login&email=&h=  shared sign-up / login accounts -> {ok:true|false, ...}
 *
 * SETUP
 *  1. Open your Google Sheet -> Extensions -> Apps Script, paste this file as Code.gs.
 *     (Or use a standalone script and keep SHEET_ID below.)
 *  2. Run setup() once and accept the permissions. It creates the tabs and headers.
 *  3. Deploy -> New deployment -> type: Web app
 *        Execute as: Me      Who has access: Anyone
 *     Copy the /exec URL and paste it as WEB_APP_URL in index.html.
 *  4. Share the Sheet as "Anyone with the link: Viewer" (the dashboard reads logs from it).
 *     Passwords are NOT stored in the Sheet; account hashes are kept in Script Properties.
 */

var SHEET_ID = '11-VVgEX5ugFP4suckiaV_LitXz2VhuYe3HhDv4rLcdA';

var ACT_TAB = 'Activity';       // must be the FIRST tab (gid=0)
var PAY_TAB = 'Payment_Log';

var ACT_HEADERS = ['Time', 'Email', 'Agent Code', 'Name', 'Action', 'Where', 'Search', 'Device', 'Visitor'];
var PAY_HEADERS = ['Saved at', 'Key', 'Subscriber', 'Account', 'Agent', 'Action', 'Label',
                   'Date of payment', 'Remarks', 'Saved by (email)', 'Agent code', 'Saved by (name)'];

/* ---------- helpers ---------- */
function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function tab_(name, headers, first) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@'); // keep text as typed
  }
  if (first) { ss.setActiveSheet(sh); ss.moveActiveSheet(1); }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function s_(v, max) {
  v = (v === null || v === undefined) ? '' : String(v);
  if (/^[=+\-@]/.test(v)) v = "'" + v;         // block spreadsheet formula injection
  return v.length > (max || 2000) ? v.slice(0, max || 2000) : v;
}

/* ---------- one-time setup ---------- */
function setup() {
  var ss = ss_();
  var first = ss.getSheets()[0];
  // make the Activity log the first tab so it is gid=0
  var act = ss.getSheetByName(ACT_TAB);
  if (!act) {
    if (first && first.getLastRow() === 0) { first.setName(ACT_TAB); act = first; }
  }
  tab_(ACT_TAB, ACT_HEADERS, true);
  tab_(PAY_TAB, PAY_HEADERS, false);
  return 'Setup complete';
}

/* ---------- web app entry points ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.acct) return json_(account_(p.acct, p.email, p.h));
    if (p.action === 'write') return json_(write_(p.data));
    if (p.action === 'read') return json_(read_(p.kind));
    return json_({ success: true, message: 'MMMDA dashboard web app is running' });
  } catch (err) {
    return json_({ success: false, ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  // optional: accept the same payload by POST
  var p = (e && e.parameter) || {};
  var data = p.data || (e && e.postData && e.postData.contents) || '';
  try { return json_(write_(data)); }
  catch (err) { return json_({ success: false, error: String(err && err.message || err) }); }
}

/* ---------- write ---------- */
function write_(raw) {
  if (!raw) throw new Error('no data');
  var o = JSON.parse(raw);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (o.type === 'pay') {
      var sh = tab_(PAY_TAB, PAY_HEADERS, false);
      sh.appendRow([
        s_(o.at), s_(o.key), s_(o.name), s_(o.account), s_(o.agent), s_(o.action), s_(o.label),
        s_(o.date), s_(o.remarks, 5000), s_(o.by), s_(o.code), s_(o.byName)
      ]);
    } else {
      var a = tab_(ACT_TAB, ACT_HEADERS, false);
      a.appendRow([
        s_(o.time), s_(o.email), s_(o.code), s_(o.name), s_(o.action),
        s_(o.where), s_(o.search, 1000), s_(o.device, 200), s_(o.visitor, 50)
      ]);
    }
  } finally {
    lock.releaseLock();
  }
  return { success: true };
}

/* ---------- read ---------- */
function read_(kind) {
  if (kind === 'activity' || kind === 'act') {
    var a = tab_(ACT_TAB, ACT_HEADERS, false);
    var v = a.getLastRow() > 1 ? a.getRange(2, 1, a.getLastRow() - 1, ACT_HEADERS.length).getDisplayValues() : [];
    v = v.slice(-5000).map(function (x) {
      return { time: x[0], email: x[1], code: x[2], name: x[3], action: x[4], where: x[5], search: x[6], device: x[7], visitor: x[8] };
    });
    return { success: true, data: v };
  }
  var sh = tab_(PAY_TAB, PAY_HEADERS, false);
  var vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, PAY_HEADERS.length).getDisplayValues() : [];
  var rows = vals.map(function (x) {
    return { at: x[0], key: x[1], name: x[2], account: x[3], agent: x[4], action: x[5], label: x[6],
             date: x[7], remarks: x[8], by: x[9], code: x[10], byName: x[11] };
  }).filter(function (r) { return r.key; });
  if (kind === 'payhist') {
    rows.sort(function (a, b) { return a.at < b.at ? 1 : -1; });
    return { success: true, data: rows.slice(0, 1000) };
  }
  var latest = {};
  rows.forEach(function (r) { if (!latest[r.key] || r.at >= latest[r.key].at) latest[r.key] = r; });
  return { success: true, data: Object.keys(latest).map(function (k) { return latest[k]; }) };
}

/* ---------- shared accounts (hashes live in Script Properties, not in the Sheet) ---------- */
function account_(op, email, h) {
  email = String(email || '').trim().toLowerCase();
  if (!email || !h) return { ok: false, reason: 'bad' };
  var props = PropertiesService.getScriptProperties();
  var key = 'acc:' + email;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var cur = props.getProperty(key);
    if (op === 'create') {
      if (cur) return { ok: false, exists: true };
      props.setProperty(key, String(h));
      return { ok: true };
    }
    if (op === 'login') {
      if (!cur) return { ok: false, reason: 'none' };
      return cur === String(h) ? { ok: true } : { ok: false, reason: 'bad' };
    }
    return { ok: false, reason: 'bad' };
  } finally {
    lock.releaseLock();
  }
}

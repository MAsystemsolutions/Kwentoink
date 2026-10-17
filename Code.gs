/******************************************************************
 *  KWENTOINK — Back end (Google Apps Script)
 *  "Tell your business story, simply."
 *
 *  What this file does:
 *   - Keeps your DEVELOPER TRACKER sheet (all customers, trial, paid-until)
 *   - Creates ONE Google Sheet + ONE receipts folder per customer
 *   - Talks to the Kwentoink app (index.html on GitHub Pages) for:
 *       loading data, saving/syncing entries, receipts, activity log
 *   - Serves your ADMIN PAGE (open on your phone to add customers,
 *     copy/send links, mark paid, reset links)
 *
 *  FIRST TIME: fill in the 3 settings below, then run setupKwentoink()
 *  once from the editor (see the setup guide).
 ******************************************************************/

// ===== 1. YOUR SETTINGS =========================================
// Your GitHub Pages address for Kwentoink (must end with a slash /)
const APP_URL = 'https://YOUR-GITHUB-USERNAME.github.io/kwentoink/';

// Shown to customers when their trial/subscription ends
const SUPPORT_CONTACT = 'Message us on Messenger or GCash 09XX XXX XXXX';

// Free trial length for new customers
const TRIAL_DAYS = 14;
// ================================================================

const OFFLINE_GRACE_DAYS = 7; // app keeps working offline this long without a check-in
const TRACKER_HEADERS = ['Code', 'Business', 'Owner', 'Contact', 'Status', 'Trial Ends', 'Paid Until',
  'Sheet ID', 'Folder ID', 'Link', 'Created', 'Notes', 'Last Seen'];
const ENTRY_COLS = ['id', 'date', 'type', 'amount', 'cat', 'catLabel', 'method', 'note', 'receiptId',
  'receiptMime', 'by', 'device', 'createdAt', 'updatedAt', 'deleted'];
const UTANG_COLS = ['id', 'kind', 'name', 'amount', 'paid', 'due', 'note', 'status', 'payments',
  'by', 'device', 'createdAt', 'updatedAt', 'deleted'];
const LOG_COLS = ['time', 'by', 'device', 'action', 'table', 'recId', 'summary'];

/* ---------------------------------------------------------------
 * ONE-TIME SETUP — run this once from the Apps Script editor
 * ------------------------------------------------------------- */
function setupKwentoink() {
  const props = PropertiesService.getScriptProperties();
  let trackerId = props.getProperty('TRACKER_ID');
  if (!trackerId) {
    const root = DriveApp.createFolder('Kwentoink (customers)');
    const ss = SpreadsheetApp.create('Kwentoink – Developer Tracker');
    DriveApp.getFileById(ss.getId()).moveTo(root);
    const sh = ss.getSheets()[0];
    sh.setName('Customers');
    sh.getRange(1, 1, 1, TRACKER_HEADERS.length).setValues([TRACKER_HEADERS])
      .setFontWeight('bold').setBackground('#cfe8fb');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    props.setProperty('TRACKER_ID', ss.getId());
    props.setProperty('ROOT_FOLDER_ID', root.getId());
    trackerId = ss.getId();
  }
  if (!props.getProperty('ADMIN_KEY')) props.setProperty('ADMIN_KEY', randomCode_(16));
  Logger.log('✅ Kwentoink is set up.');
  Logger.log('Tracker sheet: https://docs.google.com/spreadsheets/d/' + trackerId);
  Logger.log('After you Deploy as a web app, your ADMIN page is:');
  Logger.log('   <your web app URL>?admin=' + props.getProperty('ADMIN_KEY'));
  Logger.log('Keep that admin link private.');
}

/** Run this anytime from the editor if you forget your admin key. */
function showAdminKey() {
  Logger.log('Admin key: ' + PropertiesService.getScriptProperties().getProperty('ADMIN_KEY'));
  Logger.log('Admin page: <your web app URL>?admin=' + PropertiesService.getScriptProperties().getProperty('ADMIN_KEY'));
}

/* ---------------------------------------------------------------
 * WEB ENTRY POINTS
 * ------------------------------------------------------------- */
function doGet(e) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (e && e.parameter && e.parameter.admin && key && e.parameter.admin === key) {
    return HtmlService.createHtmlOutput(adminHtml_())
      .setTitle('Kwentoink Admin')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
  }
  return ContentService.createTextOutput('Kwentoink API is running 🐃');
}

function doPost(e) {
  let out;
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const fn = {
      init: api_init_, setup: api_setup_, sync: api_sync_,
      receipt: api_receipt_, getReceipt: api_getReceipt_, log: api_log_
    }[body.action];
    if (!fn) throw new Error('Unknown action');
    out = fn(body);
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------------------------------------------------------
 * APP API
 * ------------------------------------------------------------- */
function api_init_(b) {
  const c = findCustomer_(b.code);
  if (!c) return { ok: true, sub: { state: 'invalid' } };
  touchLastSeen_(c);
  return fullPayload_(c);
}

function api_setup_(b) {
  const c = requireCustomer_(b.code);
  return withLock_(() => {
    const ss = SpreadsheetApp.openById(c.sheetId);
    const s = b.settings || {};
    s.setupDone = true;
    s.setupAt = new Date().toISOString();
    writeSettings_(ss, s);
    // keep tracker business/owner name in sync
    const tr = trackerSheet_();
    if (s.businessName) tr.getRange(c.row, 2).setValue(s.businessName);
    if (s.ownerName && !c.owner) tr.getRange(c.row, 3).setValue(s.ownerName);
    appendLog_(ss, [[new Date().toISOString(), b.by || s.ownerName || 'Owner', b.device || '', 'setup', 'settings', '', s.businessName || '']]);
    return fullPayload_(c);
  });
}

function api_sync_(b) {
  const c = requireCustomer_(b.code);
  const state = subState_(c);
  if (state.state === 'expired') return fullPayload_(c); // read-only when expired
  const ops = Array.isArray(b.ops) ? b.ops : [];
  if (!ops.length) return fullPayload_(c);
  return withLock_(() => {
    const ss = SpreadsheetApp.openById(c.sheetId);
    const logs = [];
    const byTable = { entries: [], utang: [] };
    ops.forEach(op => {
      if (op.op === 'settings') {
        writeSettings_(ss, op.data || {});
        logs.push([op.at || new Date().toISOString(), op.by || '', op.device || '', 'settings', 'settings', '', op.summary || '']);
      } else if (byTable[op.table]) {
        byTable[op.table].push(op);
        logs.push([op.at || new Date().toISOString(), op.by || '', op.device || '', op.action || op.op,
          op.table, (op.rec && op.rec.id) || '', op.summary || '']);
      }
    });
    if (byTable.entries.length) upsertRows_(ss.getSheetByName('Entries'), ENTRY_COLS, byTable.entries);
    if (byTable.utang.length) upsertRows_(ss.getSheetByName('Utang'), UTANG_COLS, byTable.utang);
    if (logs.length) appendLog_(ss, logs);
    touchLastSeen_(c);
    return fullPayload_(c);
  });
}

function api_receipt_(b) {
  const c = requireCustomer_(b.code);
  if (subState_(c).state === 'expired') throw new Error('Subscription ended');
  const folder = DriveApp.getFolderById(c.folderId);
  const bytes = Utilities.base64Decode(b.data);
  const blob = Utilities.newBlob(bytes, b.mime || 'image/jpeg', (b.id || 'receipt') + (b.mime === 'application/pdf' ? '.pdf' : '.jpg'));
  const f = folder.createFile(blob);
  return { ok: true, fileId: f.getId() };
}

function api_getReceipt_(b) {
  const c = requireCustomer_(b.code);
  const f = DriveApp.getFileById(b.fileId);
  // only allow files that live in this customer's receipts folder
  const parents = f.getParents();
  let mine = false;
  while (parents.hasNext()) if (parents.next().getId() === c.folderId) mine = true;
  if (!mine) throw new Error('Not allowed');
  const blob = f.getBlob();
  return { ok: true, mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}

function api_log_(b) {
  const c = requireCustomer_(b.code);
  const sh = SpreadsheetApp.openById(c.sheetId).getSheetByName('Log');
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, log: [] };
  const n = Math.min(300, last - 1);
  const rows = sh.getRange(last - n + 1, 1, n, LOG_COLS.length).getValues().reverse();
  return { ok: true, log: rows.map(r => ({ time: str_(r[0]), by: r[1], device: r[2], action: r[3], table: r[4], recId: r[5], summary: r[6] })) };
}

/* ---------------------------------------------------------------
 * PAYLOAD / DATA
 * ------------------------------------------------------------- */
function fullPayload_(c) {
  const ss = SpreadsheetApp.openById(c.sheetId);
  const settings = readSettings_(ss);
  return {
    ok: true,
    sub: subState_(c),
    support: SUPPORT_CONTACT,
    grace: OFFLINE_GRACE_DAYS,
    settings: settings.setupDone ? settings : null,
    entries: readTable_(ss.getSheetByName('Entries'), ENTRY_COLS),
    utang: readTable_(ss.getSheetByName('Utang'), UTANG_COLS),
    serverTime: new Date().toISOString()
  };
}

function readTable_(sh, cols) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, cols.length).getValues();
  const tz = Session.getScriptTimeZone();
  return rows.filter(r => r[0]).map(r => {
    const o = {};
    cols.forEach((k, i) => {
      let v = r[i];
      if (v instanceof Date) v = (k === 'date' || k === 'due') ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : v.toISOString();
      if (k === 'amount' || k === 'paid') v = Number(v) || 0;
      if (k === 'deleted') v = v === true || v === 'TRUE' || v === 'true';
      if (k === 'payments') { try { v = v ? JSON.parse(v) : []; } catch (e) { v = []; } }
      o[k] = v === null || v === undefined ? '' : v;
    });
    return o;
  });
}

function upsertRows_(sh, cols, ops) {
  const last = sh.getLastRow();
  const data = last > 1 ? sh.getRange(2, 1, last - 1, cols.length).getValues() : [];
  const index = {};
  data.forEach((r, i) => { if (r[0]) index[r[0]] = i; });
  const updIdx = cols.indexOf('updatedAt');
  const changed = {};
  const appended = [];
  ops.forEach(op => {
    const rec = op.rec || {};
    if (!rec.id) return;
    const row = cols.map(k => {
      let v = rec[k];
      if (k === 'payments') v = JSON.stringify(v || []);
      if (k === 'deleted') v = !!v;
      return v === undefined || v === null ? '' : v;
    });
    if (index[rec.id] !== undefined) {
      const i = index[rec.id];
      const existingUpd = str_(data[i][updIdx]);
      if (existingUpd && rec.updatedAt && String(rec.updatedAt) < existingUpd) return; // older edit loses
      data[i] = row;
      changed[i] = true;
    } else {
      index[rec.id] = data.length + appended.length; // guards duplicates in same batch
      appended.push(row);
    }
  });
  Object.keys(changed).forEach(i => sh.getRange(Number(i) + 2, 1, 1, cols.length).setValues([data[i]]));
  if (appended.length) sh.getRange(sh.getLastRow() + 1, 1, appended.length, cols.length).setValues(appended);
}

function readSettings_(ss) {
  const sh = ss.getSheetByName('Settings');
  const last = sh.getLastRow();
  const o = {};
  if (last < 2) return o;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => {
    if (!r[0]) return;
    try { o[r[0]] = JSON.parse(r[1]); } catch (e) { o[r[0]] = r[1]; }
  });
  return o;
}

function writeSettings_(ss, obj) {
  const sh = ss.getSheetByName('Settings');
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
  const idx = {};
  rows.forEach((r, i) => idx[r[0]] = i);
  Object.keys(obj).forEach(k => {
    const v = JSON.stringify(obj[k]);
    if (idx[k] !== undefined) rows[idx[k]][1] = v; else { idx[k] = rows.length; rows.push([k, v]); }
  });
  if (rows.length) {
    sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  }
}

function appendLog_(ss, rows) {
  const sh = ss.getSheetByName('Log');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, LOG_COLS.length).setValues(rows);
}

/* ---------------------------------------------------------------
 * CUSTOMERS & SUBSCRIPTION
 * ------------------------------------------------------------- */
function trackerSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('TRACKER_ID');
  if (!id) throw new Error('Run setupKwentoink() first');
  return SpreadsheetApp.openById(id).getSheetByName('Customers');
}

function findCustomer_(code) {
  if (!code) return null;
  const sh = trackerSheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const rows = sh.getRange(2, 1, last - 1, TRACKER_HEADERS.length).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(code)) return rowToCustomer_(rows[i], i + 2);
  }
  return null;
}

function rowToCustomer_(r, row) {
  return {
    row: row, code: String(r[0]), business: r[1], owner: r[2], contact: r[3], status: r[4],
    trialEnds: r[5], paidUntil: r[6], sheetId: r[7], folderId: r[8], link: r[9],
    created: r[10], notes: r[11], lastSeen: r[12]
  };
}

function requireCustomer_(code) {
  const c = findCustomer_(code);
  if (!c) throw new Error('INVALID_LINK');
  return c;
}

function subState_(c) {
  const today = startOfDay_(new Date());
  const paid = c.paidUntil ? startOfDay_(new Date(c.paidUntil)) : null;
  const trial = c.trialEnds ? startOfDay_(new Date(c.trialEnds)) : null;
  const days = d => Math.round((d - today) / 86400000);
  if (String(c.status).toLowerCase() === 'paused') return { state: 'expired', reason: 'paused' };
  if (paid && paid >= today) return { state: 'active', daysLeft: days(paid), until: isoDate_(paid) };
  if (trial && trial >= today) return { state: 'trial', daysLeft: days(trial), until: isoDate_(trial) };
  return { state: 'expired' };
}

function touchLastSeen_(c) {
  try { trackerSheet_().getRange(c.row, 13).setValue(new Date()); } catch (e) { }
}

function createCustomerSheet_(business, code) {
  const root = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'));
  const folder = root.createFolder('Kwentoink – ' + business + ' (' + code + ')');
  const receipts = folder.createFolder('Receipts');
  const ss = SpreadsheetApp.create('Kwentoink – ' + business);
  DriveApp.getFileById(ss.getId()).moveTo(folder);
  const first = ss.getSheets()[0];
  first.setName('Entries');
  const make = (sh, cols, textCols) => {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#cfe8fb');
    sh.setFrozenRows(1);
    textCols.forEach(k => {
      const i = cols.indexOf(k) + 1;
      if (i > 0) sh.getRange(1, i, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  };
  make(first, ENTRY_COLS, ['id', 'date', 'createdAt', 'updatedAt', 'receiptId']);
  make(ss.insertSheet('Utang'), UTANG_COLS, ['id', 'due', 'createdAt', 'updatedAt', 'payments']);
  make(ss.insertSheet('Log'), LOG_COLS, ['time', 'recId']);
  make(ss.insertSheet('Settings'), ['key', 'value'], ['key', 'value']);
  return { sheetId: ss.getId(), folderId: receipts.getId() };
}

/* ---------------------------------------------------------------
 * ADMIN (used by your admin page)
 * ------------------------------------------------------------- */
function checkKey_(key) {
  if (!key || key !== PropertiesService.getScriptProperties().getProperty('ADMIN_KEY')) throw new Error('Wrong admin key');
}

function admin_list(key) {
  checkKey_(key);
  const sh = trackerSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, TRACKER_HEADERS.length).getValues()
    .filter(r => r[0])
    .map((r, i) => {
      const c = rowToCustomer_(r, i + 2);
      const s = subState_(c);
      return {
        code: c.code, business: c.business, owner: c.owner, contact: c.contact, status: c.status,
        state: s.state, daysLeft: s.daysLeft, until: s.until, link: c.link, notes: c.notes,
        lastSeen: c.lastSeen ? isoDate_(new Date(c.lastSeen)) : ''
      };
    }).reverse();
}

function admin_create(key, info) {
  checkKey_(key);
  return withLock_(() => {
    const business = String(info.business || 'New Business').trim();
    const code = randomCode_(10);
    const made = createCustomerSheet_(business, code);
    const trialEnds = new Date(); trialEnds.setDate(trialEnds.getDate() + TRIAL_DAYS);
    const link = APP_URL + '?id=' + code;
    trackerSheet_().appendRow([code, business, info.owner || '', info.contact || '', 'Trial', trialEnds, '',
      made.sheetId, made.folderId, link, new Date(), info.notes || '', '']);
    return { code: code, link: link };
  });
}

function admin_extend(key, code, months) {
  checkKey_(key);
  const c = requireCustomer_(code);
  const sh = trackerSheet_();
  const today = startOfDay_(new Date());
  let base = c.paidUntil ? startOfDay_(new Date(c.paidUntil)) : today;
  if (base < today) base = today;
  base.setMonth(base.getMonth() + (Number(months) || 1));
  sh.getRange(c.row, 5).setValue('Active');
  sh.getRange(c.row, 7).setValue(base);
  return admin_list(key);
}

function admin_setPaused(key, code, paused) {
  checkKey_(key);
  const c = requireCustomer_(code);
  const sh = trackerSheet_();
  const nextStatus = paused ? 'Paused' : (c.paidUntil && new Date(c.paidUntil) >= startOfDay_(new Date()) ? 'Active' : 'Trial');
  sh.getRange(c.row, 5).setValue(nextStatus);
  return admin_list(key);
}

function admin_resetLink(key, code) {
  checkKey_(key);
  const c = requireCustomer_(code);
  const newCode = randomCode_(10);
  const link = APP_URL + '?id=' + newCode;
  const sh = trackerSheet_();
  sh.getRange(c.row, 1).setValue(newCode);
  sh.getRange(c.row, 10).setValue(link);
  return admin_list(key);
}

/* ---------------------------------------------------------------
 * HELPERS
 * ------------------------------------------------------------- */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function randomCode_(n) {
  const chars = 'abcdefghjkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < n; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}
function startOfDay_(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function isoDate_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function str_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }

/* ---------------------------------------------------------------
 * ADMIN PAGE (opens at  <web app URL>?admin=YOUR_KEY )
 * ------------------------------------------------------------- */
function adminHtml_() {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return `<!doctype html><html><head><meta charset="utf-8">
<style>
:root{--sky:#cfe8fb;--navy:#2c3e6b;--pink:#ef6f8a;--cream:#fff8ee;--green:#d8f3dc;--muted:#6b7a99}
*{box-sizing:border-box}body{margin:0;font-family:ui-rounded,"SF Pro Rounded",system-ui,sans-serif;background:var(--sky);color:var(--navy)}
header{padding:18px 16px 8px;font-weight:800;font-size:22px}header span{color:var(--pink)}
.card{background:#fff;border-radius:18px;margin:12px 16px;padding:14px;box-shadow:0 2px 0 rgba(44,62,107,.08)}
input{width:100%;padding:12px;border-radius:12px;border:2px solid #e3eefa;font:inherit;margin:4px 0 8px}
button{font:inherit;border:0;border-radius:12px;padding:10px 12px;font-weight:700;cursor:pointer;background:var(--sky);color:var(--navy)}
.primary{background:var(--navy);color:#fff;width:100%;padding:14px}
.row{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px}.row button{flex:1 1 auto}
.chip{display:inline-block;padding:3px 10px;border-radius:99px;font-size:12px;font-weight:800}
.trial{background:#fff1c9}.active{background:var(--green)}.expired{background:#ffd9df}
.muted{color:var(--muted);font-size:13px}.link{word-break:break-all;font-size:12px;background:var(--cream);padding:8px;border-radius:10px;margin-top:8px}
#toast{position:fixed;left:50%;bottom:20px;transform:translateX(-50%);background:var(--navy);color:#fff;padding:10px 16px;border-radius:99px;display:none}
</style></head><body>
<header>KWENTO<span>INK</span> Admin 🐃</header>
<div class="card"><b>➕ New customer</b>
<input id="biz" placeholder="Business name"><input id="own" placeholder="Owner name (optional)">
<input id="con" placeholder="Contact: phone / FB (optional)"><input id="not" placeholder="Notes (optional)">
<button class="primary" onclick="create()">Create & get link</button></div>
<div id="list"><div class="card muted">Loading…</div></div>
<div id="toast"></div>
<script>
var KEY=${JSON.stringify(key)};
function toast(m){var t=document.getElementById('toast');t.textContent=m;t.style.display='block';setTimeout(function(){t.style.display='none'},2200)}
function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
function fail(e){toast('Error: '+(e.message||e))}
function load(){google.script.run.withSuccessHandler(render).withFailureHandler(fail).admin_list(KEY)}
function render(list){
 var h='';
 if(!list.length)h='<div class="card muted">No customers yet.</div>';
 list.forEach(function(c){
  var st=c.status==='Paused'?'Paused':(c.state==='active'?'Paid':c.state==='trial'?'Trial':'Expired');
  var info=c.until?(c.daysLeft+' days left · until '+c.until):'';
  h+='<div class="card"><div style="display:flex;justify-content:space-between;gap:8px"><b>'+esc(c.business)+'</b><span class="chip '+c.state+'">'+st+'</span></div>'
   +'<div class="muted">'+esc(c.owner||'')+(c.contact?' · '+esc(c.contact):'')+'</div>'
   +'<div class="muted">'+info+(c.lastSeen?' · last used '+c.lastSeen:'')+'</div>'
   +'<div class="link">'+esc(c.link)+'</div>'
   +'<div class="row"><button onclick="copy(\\''+c.code+'\\')">📋 Copy link</button><button onclick="sms(\\''+c.code+'\\')">💬 Text it</button></div>'
   +'<div class="row"><button onclick="ext(\\''+c.code+'\\',1)">✅ Paid +1 month</button><button onclick="pause(\\''+c.code+'\\','+(c.status!=='Paused')+')">'+(c.status==='Paused'?'▶️ Resume':'⏸ Pause')+'</button><button onclick="reset(\\''+c.code+'\\')">🔁 New link</button></div></div>';
 });
 document.getElementById('list').innerHTML=h;window._list=list;
}
function find(code){return (window._list||[]).filter(function(c){return c.code===code})[0]}
function msg(c){return 'Hi! Ito ang Kwentoink link mo 🐃 '+c.link+' — buksan sa Safari (iPhone) o Chrome (Android), tapos i-add sa Home Screen.'}
function copy(code){var c=find(code);var ta=document.createElement('textarea');ta.value=msg(c);document.body.appendChild(ta);ta.select();try{document.execCommand('copy');toast('Copied! Paste it in Messenger 💙')}catch(e){prompt('Copy this:',msg(c))}ta.remove()}
function sms(code){var c=find(code);window.open('sms:?&body='+encodeURIComponent(msg(c)),'_top')}
function create(){var b=document.getElementById('biz').value.trim();if(!b){toast('Type the business name');return}
 toast('Creating… (takes a few seconds)');
 google.script.run.withSuccessHandler(function(r){['biz','own','con','not'].forEach(function(i){document.getElementById(i).value=''});toast('Created! Link is ready 🎉');load()}).withFailureHandler(fail)
 .admin_create(KEY,{business:b,owner:document.getElementById('own').value,contact:document.getElementById('con').value,notes:document.getElementById('not').value})}
function ext(code,m){google.script.run.withSuccessHandler(function(l){render(l);toast('Marked paid +'+m+' month')}).withFailureHandler(fail).admin_extend(KEY,code,m)}
function pause(code,p){google.script.run.withSuccessHandler(render).withFailureHandler(fail).admin_setPaused(KEY,code,p)}
function reset(code){if(!confirm('Make a NEW link? The old link will stop working on all phones. Their data stays.'))return;
 google.script.run.withSuccessHandler(function(l){render(l);toast('New link made — send it to the owner')}).withFailureHandler(fail).admin_resetLink(KEY,code)}
load();
</script></body></html>`;
}

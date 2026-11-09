// ==========================================================================
// 🔑 MASTER KEY — works like the Family Key in STEAM Playground.
//
// The links that lead to this app's data (the /exec link, the Sheet ID, the
// Apps Script project ID) are locked IN THE BROWSER with the Master Key
// (PBKDF2-SHA256 150,000 rounds → AES-GCM 256). Only the locked copy and a
// hash of the key are kept: the hash in Script Properties, the locked copy
// in Drive as "Master Link.json" (next to the Sheet) and in the private
// GitHub repo as secrets/master-link.json. Whoever wants the data back —
// a backup, a deploy from the repo — must type the Master Key.
//
// First time:  run startMasterKeySetup() in the editor → it logs a one-time
//              setup code (30 min) and the page link (…/exec?page=master).
// Later:       the page asks for the current Master Key to change it.
// Data door:   POST {api:'ping'|'backup', key} to /exec (text/plain body).
// ==========================================================================

const MASTER_PROP_HASH_ = 'MASTER_KEY_HASH';
const MASTER_PROP_SETUP_ = 'MASTER_SETUP_CODE';
const MASTER_PROP_LINK_ = 'MASTER_LINK_JSON';
const MASTER_LINK_FILE_ = 'Master Link.json';
const MASTER_MAX_FAILS_ = 5;          // wrong tries …
const MASTER_LOCK_SECONDS_ = 600;     // … then locked for 10 minutes

function masterNorm_(key) { return String(key || '').trim().toLowerCase().replace(/\s+/g, ' '); }

function masterHash_(key) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'mvl-master:' + masterNorm_(key), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function masterProps_() { return PropertiesService.getScriptProperties(); }

// Wrong-key brake: after 5 wrong tries in a row, refuse everything for 10 minutes.
function masterGuard_() {
  const c = CacheService.getScriptCache();
  if (Number(c.get('mvl-master-fails') || 0) >= MASTER_MAX_FAILS_) {
    throw new Error('🔒 Too many wrong tries — wait 10 minutes and try again.');
  }
}
function masterFail_() {
  const c = CacheService.getScriptCache();
  c.put('mvl-master-fails', String(Number(c.get('mvl-master-fails') || 0) + 1), MASTER_LOCK_SECONDS_);
}
function masterOk_() { CacheService.getScriptCache().remove('mvl-master-fails'); }

function masterKeyIsSet_() { return !!masterProps_().getProperty(MASTER_PROP_HASH_); }

function masterCheckKey_(key) {
  masterGuard_();
  const stored = masterProps_().getProperty(MASTER_PROP_HASH_);
  if (!stored || !key || masterHash_(key) !== stored) { masterFail_(); return false; }
  masterOk_();
  return true;
}

function masterCheckSetupCode_(code) {
  masterGuard_();
  const raw = masterProps_().getProperty(MASTER_PROP_SETUP_) || '';
  const parts = raw.split('|');
  const ok = parts.length === 2 && Date.now() < Number(parts[1]) &&
    String(code || '').trim().toUpperCase() === parts[0];
  if (!ok) { masterFail_(); return false; }
  masterOk_();
  return true;
}

// proof = { key } (Master Key already set) or { code } (first time only).
function masterProve_(proof) {
  proof = proof || {};
  if (masterKeyIsSet_()) {
    if (!masterCheckKey_(proof.key)) throw new Error('🔑 That Master Key is not right.');
  } else if (!masterCheckSetupCode_(proof.code)) {
    throw new Error('🔑 Setup code is wrong or expired — run startMasterKeySetup() in the editor again.');
  }
}

// ---- Run from the editor (first time, or if the Master Key is forgotten) ----
function startMasterKeySetup() {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 8; i++) code += abc.charAt(Math.floor(Math.random() * abc.length));
  masterProps_().setProperty(MASTER_PROP_SETUP_, code + '|' + (Date.now() + 30 * 60 * 1000));
  // Forgotten key: the editor (= the Google account owner) may reset it.
  masterProps_().deleteProperty(MASTER_PROP_HASH_);
  CacheService.getScriptCache().remove('mvl-master-fails');
  DriveApp.getRootFolder(); // asks for Drive permission now, so the page can save "Master Link.json"
  const url = ScriptApp.getService().getUrl();
  Logger.log('🔑 Setup code (valid 30 minutes): ' + code);
  Logger.log('Open: ' + (url ? url + '?page=master' : 'your /exec link + ?page=master'));
  return code;
}

// ---- Called by master.html ----
function masterStatus() {
  return { set: masterKeyIsSet_() };
}

// Returns the private links ONLY after the code / key is proven.
function masterUnlock(proof) {
  masterProve_(proof);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    execUrl: ScriptApp.getService().getUrl() || '',
    sheetId: ss.getId(),
    sheetName: ss.getName(),
    scriptId: ScriptApp.getScriptId(),
  };
}

// newKey: 3+ words. link: the locked copy made in the browser {v,salt,iv,ct,iter}.
function saveMasterKey(proof, newKey, link) {
  masterProve_(proof);
  const words = masterNorm_(newKey).split(' ').filter(Boolean);
  if (words.length < 3) throw new Error('✋ Pick 3 or more words — easy to type, hard to guess.');
  if (!link || !link.ct || !link.salt || !link.iv) throw new Error('The locked link is missing — try again.');
  const saved = { v: 1, salt: String(link.salt), iv: String(link.iv), ct: String(link.ct), iter: Number(link.iter) || 150000, at: new Date().toISOString() };
  const json = JSON.stringify(saved);
  const props = masterProps_();
  props.setProperty(MASTER_PROP_HASH_, masterHash_(newKey));
  props.setProperty(MASTER_PROP_LINK_, json);
  props.deleteProperty(MASTER_PROP_SETUP_);
  let drive = false;
  try { drive = masterWriteLinkFile_(json); } catch (e) { Logger.log('Master Link.json not written: ' + e); }
  return { ok: true, drive: drive };
}

// Writes "Master Link.json" into the Sheet's Drive folder (replaces the old one).
function masterWriteLinkFile_(json) {
  const sheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const parents = sheetFile.getParents();
  const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const old = folder.getFilesByName(MASTER_LINK_FILE_);
  while (old.hasNext()) old.next().setTrashed(true);
  folder.createFile(MASTER_LINK_FILE_, json, MimeType.PLAIN_TEXT);
  return true;
}

// ---- Data door: only answers with the right Master Key ----
function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { body = {}; }
  let out;
  try {
    if (!masterKeyIsSet_()) throw new Error('Master Key is not set up.');
    if (!masterCheckKey_(body.key)) throw new Error('🔑 That Master Key is not right.');
    if (body.api === 'ping') out = { ok: true, app: 'math-vs-lottery', at: new Date().toISOString() };
    else if (body.api === 'backup') out = masterBackup_();
    else throw new Error('Unknown request.');
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// Every tab of the Sheet as rows (dates as ISO text) — a full copy of the data.
function masterBackup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tabs = {};
  ss.getSheets().forEach(function (sh) {
    const r = sh.getLastRow(), c = sh.getLastColumn();
    tabs[sh.getName()] = r && c ? sh.getRange(1, 1, r, c).getValues().map(function (row) {
      return row.map(function (v) { return v instanceof Date ? v.toISOString() : v; });
    }) : [];
  });
  return { ok: true, app: 'math-vs-lottery', sheetName: ss.getName(), at: new Date().toISOString(), tabs: tabs };
}

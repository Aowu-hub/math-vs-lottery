// npm test — runs the real src/MasterKey.gs with fake Google services and
// checks the whole Master Key flow end to end (no Google account needed).
import { readFileSync } from 'node:fs';
import { createContext, runInContext } from 'node:vm';
import { createHash } from 'node:crypto';
import assert from 'node:assert/strict';
import { ROOT, lockLink, openLink } from './lib.mjs';

function makeGas() {
  const props = {}, cache = {}, drive = [], logs = [];
  const sheet = (name, rows) => ({ getName: () => name, getLastRow: () => rows.length, getLastColumn: () => rows[0] ? rows[0].length : 0,
    getRange: () => ({ getValues: () => rows }) });
  const tabs = [sheet('Historical Draws', [['Draw Date', 'Jackpot'], [new Date('2026-11-01T00:00:00Z'), '123456']]), sheet('Config', [['Key', 'Value']])];
  const ss = { getId: () => 'SHEET_ID_FOR_TEST_1234567890', getName: () => 'LOTTERY JACKPOT', getSheets: () => tabs };
  const folder = { getFilesByName: n => { const l = drive.filter(f => f.name === n && !f.trashed); let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; },
    createFile: (name, content) => { const f = { name, content, trashed: false, setTrashed(v) { this.trashed = v; } }; drive.push(f); return f; } };
  return {
    props, cache, drive, logs,
    ctx: {
      PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; } }) },
      CacheService: { getScriptCache: () => ({ get: k => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = String(v); }, remove: k => { delete cache[k]; } }) },
      Utilities: { DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
        computeDigest: (alg, text) => Array.from(createHash('sha256').update(text, 'utf8').digest()).map(b => (b > 127 ? b - 256 : b)) },
      SpreadsheetApp: { getActiveSpreadsheet: () => ss },
      DriveApp: { getRootFolder: () => folder, getFileById: () => ({ getParents: () => { let d = false; return { hasNext: () => !d, next: () => { d = true; return folder; } }; } }) },
      ScriptApp: { getService: () => ({ getUrl: () => 'https://example.invalid/exec' }), getScriptId: () => 'SCRIPT_ID_FOR_TEST_1234567890' },
      ContentService: { MimeType: { JSON: 'json' }, createTextOutput: t => ({ text: t, setMimeType() { return this; } }) },
      MimeType: { PLAIN_TEXT: 'text/plain' },
      Logger: { log: m => logs.push(String(m)) },
      Date, Math, JSON, String, Number, Error,
    },
  };
}

const gas = makeGas();
const g = createContext(gas.ctx);
runInContext(readFileSync(ROOT + 'src/MasterKey.gs', 'utf8'), g, { filename: 'MasterKey.gs' });
const call = (fn, ...a) => g[fn](...a);
const post = body => JSON.parse(call('doPost', { postData: { contents: JSON.stringify(body) } }).text);
let n = 0; const ok = m => console.log('  ✓ ' + m + (++n ? '' : ''));

// 1. nothing set yet
assert.equal(call('masterStatus').set, false); ok('starts with no Master Key');
assert.equal(post({ api: 'ping', key: 'anything' }).ok, false); ok('data door closed before setup');

// 2. setup code from the editor
const code = call('startMasterKeySetup');
assert.match(code, /^[A-Z2-9]{8}$/); assert.ok(gas.logs.some(l => l.includes(code))); ok('setup code made and logged');
assert.throws(() => call('masterUnlock', { code: 'WRONG123' })); ok('wrong setup code refused');
const info = call('masterUnlock', { code });
assert.equal(info.sheetId, 'SHEET_ID_FOR_TEST_1234567890'); ok('right setup code unlocks the private IDs');

// 3. lock + save the key
const KEY = 'Blue Moon  River Seven';
const data = { u: 'https://script.google.com/macros/s/TESTDEPLOY/exec', deploymentId: 'TESTDEPLOY', sheetId: info.sheetId, sheetName: info.sheetName, scriptId: info.scriptId };
const link = await lockLink(KEY, data);
assert.throws(() => call('saveMasterKey', { code }, 'two words', link)); ok('key with < 3 words refused');
const r = call('saveMasterKey', { code }, KEY, link);
assert.equal(r.ok, true); assert.equal(r.drive, true); ok('Master Key saved, "Master Link.json" written to Drive');
assert.equal(call('masterStatus').set, true);
assert.ok(!JSON.stringify(gas.props).includes('blue moon')); ok('only a hash of the key is stored');
assert.equal(gas.props.MASTER_SETUP_CODE, undefined); ok('setup code is used up');
assert.throws(() => call('masterUnlock', { code })); ok('setup code cannot be reused');

// 4. the locked file opens with the key only
const file = JSON.parse(gas.drive.find(f => f.name === 'Master Link.json' && !f.trashed).content);
assert.deepEqual(await openLink('blue moon river seven', file), data); ok('locked file opens with the key (any case/spaces)');
await assert.rejects(openLink('blue moon river eight', file)); ok('locked file stays shut with a wrong key');

// 5. data door
assert.equal(post({ api: 'ping', key: KEY }).ok, true); ok('ping with the key');
const b = post({ api: 'backup', key: KEY });
assert.equal(b.ok, true); assert.equal(b.tabs['Historical Draws'][1][0], '2026-11-01T00:00:00.000Z'); ok('backup returns every tab (dates as text)');
assert.equal(post({ api: 'backup', key: 'nope' }).ok, false); ok('backup refused with a wrong key');

// 6. change key: needs the current key
const link2 = await lockLink('green tea morning walk', data);
assert.throws(() => call('saveMasterKey', { key: 'wrong key here' }, 'green tea morning walk', link2)); ok('changing needs the current key');
call('saveMasterKey', { key: KEY }, 'green tea morning walk', link2);
assert.equal(post({ api: 'ping', key: KEY }).ok, false); assert.equal(post({ api: 'ping', key: 'green tea morning walk' }).ok, true); ok('new key works, old key stops working');
assert.equal(gas.drive.filter(f => f.name === 'Master Link.json' && !f.trashed).length, 1); ok('old Drive file replaced');

// 7. wrong-key brake
for (let i = 0; i < 5; i++) post({ api: 'ping', key: 'bad ' + i });
assert.match(post({ api: 'ping', key: 'green tea morning walk' }).error, /Too many/); ok('5 wrong tries lock the door for 10 minutes');

// 8. website door (no key) — only the page's functions
runInContext('function generateSet() { return { jackpot: "123456" }; } function setupSheets() { return "danger"; }', g);
const w = post({ api: 'run', fn: 'generateSet', args: [] });
assert.equal(w.ok, true); assert.equal(w.result.jackpot, '123456'); ok('website door runs a page function without a key');
assert.equal(post({ api: 'run', fn: 'setupSheets' }).ok, false); ok('website door refuses functions not on the list');
assert.equal(post({ api: 'run', fn: 'masterBackup_' }).ok, false); ok('website door cannot reach the backup');
assert.equal(post({ api: 'run', fn: 'getMathBreakdown' }).ok, false); ok('missing function answers with an error, not a crash');

console.log('✅ Master Key: ' + n + ' checks passed');

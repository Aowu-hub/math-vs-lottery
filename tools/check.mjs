// npm run check — runs before every deploy and on GitHub for every push.
// Catches the mistakes that cost us time before:
//  • a script that doesn't parse (Apps Script then silently stops the page)
//  • a one-line <script> ending in a // comment (breaks inside HtmlService)
//  • the same function name in two .gs files (Apps Script keeps only one)
//  • private links / IDs written into the code
import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { Script } from 'node:vm';
import { ROOT, LOCAL_FILE } from './lib.mjs';

const problems = [];
const src = ROOT + 'src/';
const files = readdirSync(src);

// 1) .gs files parse, and no duplicate top-level functions
const seen = {};
files.filter(f => f.endsWith('.gs')).forEach(f => {
  const code = readFileSync(src + f, 'utf8');
  try { new Script(code, { filename: f }); } catch (e) { problems.push(f + ': ' + e.message); }
  for (const m of code.matchAll(/^function\s+([\w$]+)\s*\(/gm)) {
    if (seen[m[1]]) problems.push('function ' + m[1] + '() is in both ' + seen[m[1]] + ' and ' + f);
    seen[m[1]] = f;
  }
});

// 2) every inline <script> in the .html files parses; no one-line script with a // comment
files.filter(f => f.endsWith('.html')).forEach(f => {
  const html = readFileSync(src + f, 'utf8');
  let n = 0;
  for (const m of html.matchAll(/<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)<\/script>/g)) {
    n++;
    const body = m[1];
    const line = html.slice(0, m.index).split('\n').length;
    if (!body.includes('\n') && /\/\/[^\n]*$/.test(body)) problems.push(f + ':' + line + ' one-line <script> ends with a // comment (breaks in Apps Script)');
    try { new Script(body, { filename: f + ':' + line }); } catch (e) { problems.push(f + ':' + line + ' script #' + n + ': ' + e.message); }
  }
  if (/<\?/.test(html) && !/createTemplateFromFile/.test(readFileSync(src + 'Code.gs', 'utf8'))) {
    problems.push(f + ' contains "<?" but pages are not templates — it will show as text');
  }
});

// 3) nothing private in the code (the repo is private, but code gets copied/shared)
const patterns = [
  [/docs\.google\.com\/spreadsheets\/d\/[\w-]{20,}/, 'a Google Sheet link'],
  [/script\.google\.com\/macros\/s\/[\w-]{20,}/, 'an Apps Script /exec link'],
  [/[\w.+-]+@gmail\.com/, 'an email address'],
  [/AIza[\w-]{30,}/, 'a Google API key'],
  [/(openById|getFolderById|getFileById)\(\s*['"][\w-]{20,}/, 'a hard-coded Google file/folder ID (use Script Properties)'],
  [/\b(sk-[\w-]{20,}|Bearer\s+[\w.-]{20,})/, 'an API token'],
  [/api\.line\.me[\s\S]{0,200}['"][A-Za-z0-9+/=]{100,}['"]/, 'a LINE channel token'],
];
let secrets = [];
if (existsSync(LOCAL_FILE)) {
  const s = JSON.parse(readFileSync(LOCAL_FILE, 'utf8'));
  secrets = [s.sheetId, s.scriptId, s.deploymentId].filter(v => v && v.length > 10);
}
const scan = dir => readdirSync(ROOT + dir, { withFileTypes: true }).forEach(d => {
  const rel = dir + d.name;
  if (d.isDirectory()) { if (!['.git', 'node_modules', 'backups', '.local', 'secrets', 'site'].includes(d.name)) scan(rel + '/'); return; }
  if (['.clasp.json'].includes(d.name) || /\.(png|jpg|xlsx)$/i.test(d.name) || rel.startsWith('tools/check.mjs')) return;
  const text = readFileSync(ROOT + rel, 'utf8');
  patterns.forEach(([re, what]) => { if (re.test(text)) problems.push(rel + ' contains ' + what); });
  secrets.forEach(v => { if (text.includes(v)) problems.push(rel + ' contains a private ID from the unlocked secrets'); });
});
scan('');

if (problems.length) {
  console.error('❌ ' + problems.length + ' problem(s):\n  • ' + problems.join('\n  • '));
  process.exit(1);
}
console.log('✅ Checks passed: ' + files.length + ' source files' + (secrets.length ? ' (+ private-ID scan)' : ''));

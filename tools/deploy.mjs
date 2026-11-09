// npm run deploy -- "what changed"   (also run by GitHub on every update)
//  1. checks + Master Key test
//  2. SAFETY: downloads the live Apps Script project into a temp folder and
//     stops if it has a file we don't have, or different settings
//     (appsscript.json) — so a deploy can never delete or change something
//     that was edited in the browser editor
//  3. pushes src/ and updates the SAME /exec deployment (users keep the link)
import { readFileSync, existsSync, mkdtempSync, writeFileSync, readdirSync, rmSync } from 'node:fs';
import { execSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { ROOT, LOCAL_FILE } from './lib.mjs';

if (!existsSync(LOCAL_FILE)) { console.error('🔒 Run "npm run unlock" first.'); process.exit(1); }
const { scriptId, deploymentId } = JSON.parse(readFileSync(LOCAL_FILE, 'utf8'));
// The note goes into a shell command, so keep only safe characters.
const note = (process.argv.slice(2).join(' ') || 'update').split(/\r?\n/)[0]
  .replace(/[^\p{L}\p{N}\p{M} .,:#()\/+_-]/gu, ' ').replace(/\s+/g, ' ').trim().slice(0, 90) || 'update';
const BIN = join(ROOT, 'node_modules', '.bin', process.platform === 'win32' ? 'clasp.cmd' : 'clasp');
const CLASP = existsSync(BIN) ? '"' + BIN + '"' : 'npx --no-install clasp';
const run = (c, opts = {}) => {
  console.log('> ' + c.split(deploymentId).join('<deployment>').split(scriptId).join('<script>'));
  return execSync(c, { stdio: opts.quiet ? 'pipe' : 'inherit', cwd: opts.cwd || ROOT, encoding: 'utf8' });
};

run('node tools/check.mjs');
run('node tools/test-masterkey.mjs');

// ---- safety check against the live project ----
const tmp = mkdtempSync(join(tmpdir(), 'mvl-live-'));
writeFileSync(join(tmp, '.clasp.json'), JSON.stringify({ scriptId, rootDir: '.' }));
run(CLASP + ' pull', { cwd: tmp, quiet: true });
const base = f => f.replace(/\.(gs|js|html|json)$/, '');
const live = readdirSync(tmp).filter(f => f !== '.clasp.json');
const ours = readdirSync(ROOT + 'src');
const missing = live.filter(f => !ours.some(o => base(o) === base(f)));
const norm = s => JSON.stringify(JSON.parse(s));
const liveManifest = live.includes('appsscript.json') ? readFileSync(join(tmp, 'appsscript.json'), 'utf8') : null;
const ourManifest = readFileSync(ROOT + 'src/appsscript.json', 'utf8');
const problems = [];
if (missing.length) problems.push('The live project has file(s) not in the repo: ' + missing.join(', ') + ' — add them to src/ first (a deploy would delete them).');
if (liveManifest && norm(liveManifest) !== norm(ourManifest)) {
  problems.push('Live appsscript.json differs from src/appsscript.json — copy the live one into the repo first:\n' + liveManifest);
}
rmSync(tmp, { recursive: true, force: true });
if (problems.length) { console.error('🛑 Deploy stopped (nothing was changed):\n  • ' + problems.join('\n  • ')); process.exit(2); }
console.log('✅ Live project matches the repo layout — safe to deploy.');

run(CLASP + ' push --force');
run(CLASP + ' deploy -i ' + deploymentId + ' -d "' + note + '"');
console.log('✅ Live on the same /exec link: ' + note);

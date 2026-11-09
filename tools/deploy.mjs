// npm run deploy -- "what changed" — runs the checks, pushes src/ to Apps
// Script, and updates the SAME /exec deployment (users keep the same link).
import { readFileSync, existsSync } from 'node:fs';
import { execSync } from 'node:child_process';
import { LOCAL_FILE } from './lib.mjs';

if (!existsSync(LOCAL_FILE)) { console.error('🔒 Run "npm run unlock" first.'); process.exit(1); }
const { deploymentId } = JSON.parse(readFileSync(LOCAL_FILE, 'utf8'));
const note = process.argv.slice(2).join(' ') || 'update';
const run = c => { console.log('> ' + c.replace(deploymentId, '<deployment>')); execSync(c, { stdio: 'inherit' }); };
run('node tools/check.mjs');
run('npx clasp push --force');
run('npx clasp deploy -i ' + deploymentId + ' -d ' + JSON.stringify(note));
console.log('✅ Live on the same /exec link. Hard-refresh (Ctrl+F5) to see it.');

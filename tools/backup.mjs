// npm run backup — asks for the Master Key, opens the locked /exec link and
// downloads a full copy of every Sheet tab to backups/ (git-ignored).
import { writeFileSync, mkdirSync } from 'node:fs';
import { ROOT, readLink, openLink, askKey } from './lib.mjs';

const key = await askKey();
const info = await openLink(key, readLink());
const res = await fetch(info.u, { method: 'POST', redirect: 'follow', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify({ api: 'backup', key }) });
const data = await res.json();
if (!data.ok) { console.error('❌ ' + data.error); process.exit(1); }
mkdirSync(ROOT + 'backups', { recursive: true });
const file = ROOT + 'backups/' + data.at.replace(/[:.]/g, '-') + '.json';
writeFileSync(file, JSON.stringify(data, null, 1));
console.log('✅ Backup saved: ' + file + ' (' + Object.keys(data.tabs).length + ' tabs)');

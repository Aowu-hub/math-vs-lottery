// npm run unlock — type the Master Key once; writes the files clasp needs
// (.clasp.json and .local/secrets.json). Both are git-ignored: they never
// go into the repo.
import { writeFileSync, mkdirSync } from 'node:fs';
import { ROOT, LOCAL_FILE, readLink, openLink, askKey } from './lib.mjs';

const info = await openLink(await askKey(), readLink());
// On GitHub: hide the IDs in the run log.
if (process.env.GITHUB_ACTIONS) [info.scriptId, info.deploymentId, info.sheetId, info.u].forEach(v => v && console.log('::add-mask::' + v));
mkdirSync(ROOT + '.local', { recursive: true });
writeFileSync(LOCAL_FILE, JSON.stringify(info, null, 2));
writeFileSync(ROOT + '.clasp.json', JSON.stringify({ scriptId: info.scriptId, rootDir: 'src' }, null, 2));
console.log('✅ Unlocked: ' + info.sheetName + '. You can now run: npm run push / npm run deploy / npm run backup');

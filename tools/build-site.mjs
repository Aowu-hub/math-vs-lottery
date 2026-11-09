// npm run build:site — turns src/index.html into the public website in site/
// (GitHub Pages). Same page, plus: google.script shim → /exec website door,
// self-hosted Thai fonts, app icon + manifest (Add to Home Screen),
// service worker (opens instantly, update card). Output is stable: same
// input → same files; the version is a fingerprint of the content.
import { readFileSync, writeFileSync, mkdirSync, rmSync, copyFileSync, existsSync, readdirSync, statSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { join, relative } from 'node:path';
import { ROOT, LOCAL_FILE } from './lib.mjs';

const OUT = join(ROOT, 'site');
const WEB = join(ROOT, 'web');
const sha = b => createHash('sha256').update(b).digest('hex').slice(0, 16);

// ---- the /exec link (public website door) ----
let exec = process.env.MVL_EXEC_URL || '';
if (!exec && existsSync(LOCAL_FILE)) exec = JSON.parse(readFileSync(LOCAL_FILE, 'utf8')).u || '';
if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(exec) && !process.env.MVL_TEST_EXEC) {
  console.error('🔒 No /exec link: run "npm run unlock" first (or set MVL_EXEC_URL).'); process.exit(1);
}
if (process.env.MVL_TEST_EXEC) exec = process.env.MVL_TEST_EXEC;

rmSync(OUT, { recursive: true, force: true });
mkdirSync(join(OUT, 'fonts'), { recursive: true });
mkdirSync(join(OUT, 'icons'), { recursive: true });

// ---- self-hosted fonts (Thai + Latin only, the weights the page uses) ----
const FONTS = { mitr: [400, 500, 600], sarabun: [400, 500, 600, 700, 800] };
let fontCss = '/* Self-hosted Mitr + Sarabun (SIL Open Font License) */\n';
for (const [fam, weights] of Object.entries(FONTS)) {
  const pkg = join(ROOT, 'node_modules', '@fontsource', fam);
  for (const w of weights) {
    const css = readFileSync(join(pkg, w + '.css'), 'utf8');
    for (const block of css.split('/* ').slice(1)) {
      const name = block.slice(0, block.indexOf(' */'));
      if (!new RegExp('^' + fam + '-(thai|latin)-' + w + '-normal$').test(name)) continue;
      const file = name + '.woff2';
      copyFileSync(join(pkg, 'files', file), join(OUT, 'fonts', file));
      const range = (block.match(/unicode-range:\s*([^;]+);/) || [])[1];
      fontCss += '@font-face{font-family:\'' + fam[0].toUpperCase() + fam.slice(1) + '\';font-style:normal;font-display:swap;font-weight:' + w +
        ';src:url(' + file + ') format(\'woff2\');unicode-range:' + range + '}\n';
    }
  }
}
writeFileSync(join(OUT, 'fonts', 'fonts.css'), fontCss);

// ---- page ----
let html = readFileSync(join(ROOT, 'src', 'index.html'), 'utf8');
const fontLinks = /\s*<link rel="preconnect" href="https:\/\/fonts\.googleapis\.com">\s*<link rel="preconnect" href="https:\/\/fonts\.gstatic\.com" crossorigin>\s*<link href="https:\/\/fonts\.googleapis\.com\/css2[^"]*" rel="stylesheet">/;
if (!fontLinks.test(html)) { console.error('🛑 Could not find the Google Fonts links in src/index.html — update build-site.mjs.'); process.exit(1); }
html = html.replace(fontLinks, '\n  <link rel="stylesheet" href="fonts/fonts.css">');
const headAdd = [
  '<script src="config.js"></script>',
  '<script src="gas-shim.js"></script>',
  '<link rel="manifest" href="manifest.webmanifest">',
  '<link rel="icon" type="image/png" sizes="192x192" href="icons/icon-192.png">',
  '<link rel="apple-touch-icon" href="icons/icon-180.png">',
  '<meta name="apple-mobile-web-app-capable" content="yes">',
  '<meta name="mobile-web-app-capable" content="yes">',
  '<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">',
  '<meta name="apple-mobile-web-app-title" content="Math Lotto">',
  '<meta name="description" content="คณิตศาสตร์ ปะทะ หวย — Thai lottery statistics for fun">',
  '<title>Math vs. Lottery</title>',
].join('\n  ');
if (!html.includes('<base target="_top">')) { console.error('🛑 <base target="_top"> not found — update build-site.mjs.'); process.exit(1); }
html = html.replace('<base target="_top">', headAdd);
if (/<\?/.test(html)) { console.error('🛑 Apps Script scriptlet <? found in the page — cannot build.'); process.exit(1); }
html = html.replace('</body>', '  <script src="sw-register.js"></script>\n</body>');
writeFileSync(join(OUT, 'index.html'), html);

writeFileSync(join(OUT, 'config.js'), 'window.MVL_EXEC = ' + JSON.stringify(exec) + ';\n');
copyFileSync(join(WEB, 'gas-shim.js'), join(OUT, 'gas-shim.js'));
copyFileSync(join(WEB, 'sw-register.js'), join(OUT, 'sw-register.js'));
writeFileSync(join(OUT, '.nojekyll'), '');

// ---- icons (made by tools/make-icons.mjs from web/icon.svg) ----
for (const f of readdirSync(join(WEB, 'icons'))) copyFileSync(join(WEB, 'icons', f), join(OUT, 'icons', f));
writeFileSync(join(OUT, 'manifest.webmanifest'), JSON.stringify({
  name: 'Math vs. Lottery — คณิตศาสตร์ ปะทะ หวย',
  short_name: 'Math Lotto',
  description: 'Thai lottery statistics for fun',
  lang: 'th',
  start_url: './',
  scope: './',
  display: 'standalone',
  orientation: 'portrait',
  background_color: '#0b0a24',
  theme_color: '#0b0a24',
  icons: [
    { src: 'icons/icon-192.png', sizes: '192x192', type: 'image/png' },
    { src: 'icons/icon-512.png', sizes: '512x512', type: 'image/png' },
    { src: 'icons/icon-maskable-512.png', sizes: '512x512', type: 'image/png', purpose: 'maskable' },
  ],
}, null, 2));

// ---- privacy scan (stop on anything private) ----
const all = [];
const walk = d => readdirSync(d).forEach(f => { const p = join(d, f); statSync(p).isDirectory() ? walk(p) : all.push(p); });
walk(OUT);
const secrets = existsSync(LOCAL_FILE) ? (({ sheetId, scriptId }) => [sheetId, scriptId].filter(Boolean))(JSON.parse(readFileSync(LOCAL_FILE, 'utf8'))) : [];
const bad = [];
for (const p of all) {
  if (/\.(woff2|png)$/.test(p)) continue;
  const t = readFileSync(p, 'utf8'), rel = relative(OUT, p);
  if (/docs\.google\.com\/spreadsheets\/d\//.test(t)) bad.push(rel + ': Sheet link');
  if (/[\w.+-]+@gmail\.com/.test(t)) bad.push(rel + ': email');
  if (/AIza[\w-]{30,}/.test(t)) bad.push(rel + ': API key');
  const execs = t.match(/script\.google\.com\/macros\/s\/[\w-]+/g) || [];
  if (execs.length && rel !== 'config.js') bad.push(rel + ': an /exec link outside config.js');
  secrets.forEach(s => { if (t.includes(s)) bad.push(rel + ': private ID'); });
}
if (bad.length) { console.error('🛑 Privacy check failed:\n  • ' + bad.join('\n  • ')); rmSync(OUT, { recursive: true, force: true }); process.exit(1); }

// ---- service worker + files.json (version = content fingerprint) ----
const files = {};
all.filter(p => !p.endsWith('.nojekyll')).map(p => relative(OUT, p).split('\\').join('/')).sort()
  .forEach(rel => { files[rel] = sha(readFileSync(join(OUT, rel))); });
const version = 'v' + sha(JSON.stringify(files)).slice(0, 10);
writeFileSync(join(OUT, 'files.json'), JSON.stringify({ version, files }, null, 1));
writeFileSync(join(OUT, 'sw.js'), readFileSync(join(WEB, 'sw.template.js'), 'utf8')
  .replace('__VERSION__', version).replace('__FILES__', JSON.stringify(Object.assign({ './': files['index.html'] }, files))));
console.log('✅ Website built: ' + Object.keys(files).length + ' files, version ' + version + ' → site/');

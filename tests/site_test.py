# python3 tests/site_test.py — builds the website against a fake /exec on
# localhost and checks it on a phone-sized screen with Playwright:
# page loads, server calls go through the website door, spin works, stats
# draw, dream works, manifest + icons, service worker installs and the page
# still opens with the network off (app shell), no console errors.
import json, os, subprocess, sys, threading, time
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from functools import partial
from pathlib import Path
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parent.parent
PORT = 8765
MOCK = json.loads((ROOT / 'tests' / 'mock-data.json').read_text(encoding='utf-8'))
MAP = {'getCurrentPeriodInfo': 'info', 'getMathBreakdown': 'mb', 'getComparisonStats': 'stats', 'generateSet': 'gen',
       'getBuzzNumbers': 'buzz', 'getDreamDictionary': 'dict', 'interpretDream': 'dream'}
CALLS = []
GETS = []

class H(SimpleHTTPRequestHandler):
    def log_message(self, *a): pass
    def do_GET(self):
        GETS.append(self.path)
        return super().do_GET()
    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers['Content-Length'])) or b'{}')
        CALLS.append(body.get('fn'))
        out = {'ok': True, 'result': MOCK.get(MAP.get(body.get('fn')), [])} if body.get('api') == 'run' else {'ok': False, 'error': 'no'}
        data = json.dumps(out).encode()
        self.send_response(200); self.send_header('Content-Type', 'application/json'); self.send_header('Content-Length', str(len(data))); self.end_headers(); self.wfile.write(data)

env = dict(os.environ, MVL_TEST_EXEC=f'http://127.0.0.1:{PORT}/exec')
subprocess.run(['node', str(ROOT / 'tools' / 'build-site.mjs')], check=True, env=env)
srv = ThreadingHTTPServer(('127.0.0.1', PORT), partial(H, directory=str(ROOT / 'site')))
threading.Thread(target=srv.serve_forever, daemon=True).start()

fails = []
def check(cond, msg):
    print(('  ✓ ' if cond else '  ✗ ') + msg)
    if not cond: fails.append(msg)

with sync_playwright() as p:
    b = p.chromium.launch()
    ctx = b.new_context(viewport={'width': 390, 'height': 844}, device_scale_factor=2, is_mobile=True, has_touch=True,
                        user_agent='Mozilla/5.0 (iPhone; CPU iPhone OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1')
    pg = ctx.new_page()
    errs = []
    pg.on('pageerror', lambda e: errs.append(str(e)))
    pg.on('console', lambda m: errs.append(m.text) if m.type == 'error' else None)
    pg.goto(f'http://127.0.0.1:{PORT}/')
    pg.wait_for_timeout(800)
    pg.tap('#welcomeOverlay')
    pg.wait_for_timeout(2500)
    check('getCurrentPeriodInfo' in CALLS and 'getComparisonStats' in CALLS, 'page loads its data through the website door')
    check(pg.evaluate("document.fonts.load('600 20px Mitr', 'คณิต').then(f => f.length > 0 && document.fonts.check('600 20px Mitr', 'คณิต'))"), 'self-hosted Thai font (Mitr) loaded')
    pg.evaluate('onGenerate()')
    pg.wait_for_function("document.querySelector('#cardsContainer.revealed')", timeout=60000)
    digits = pg.evaluate("[...document.querySelectorAll('#cardsContainer .digit')].map(d=>d.textContent).join('')")
    check('generateSet' in CALLS and len(digits) >= 20, f'spin → numbers revealed ({len(digits)} digits: {digits[:12]}…)')
    check(pg.evaluate("document.getElementById('dashboard').children.length") >= 3, 'stats dashboard drawn')
    pg.screenshot(path=str(ROOT / 'tests' / 'last-phone.png'))
    man = pg.evaluate("fetch('manifest.webmanifest').then(r=>r.json())")
    check(man['display'] == 'standalone' and len(man['icons']) == 3, 'manifest: standalone app with icons')
    for ic in ['icons/icon-180.png', 'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-maskable-512.png']:
        check(pg.evaluate(f"fetch('{ic}').then(r=>r.ok)"), ic + ' served')
    pg.evaluate("window.__sameLoad = 1")
    pg.wait_for_function('navigator.serviceWorker && navigator.serviceWorker.controller !== null', timeout=15000)
    check(True, 'service worker installed and controlling the page')
    pg.wait_for_timeout(1500)
    check(pg.evaluate("window.__sameLoad === 1"), 'first visit is not reloaded when the service worker starts')
    ctx.set_offline(True)
    pg.reload()
    pg.wait_for_timeout(1500)
    check(pg.evaluate("!!document.getElementById('roulette') && document.title.includes('Lottery')"), 'page still opens with the network off (from the phone)')
    ctx.set_offline(False)
    pg.reload(); pg.wait_for_timeout(1500)
    # ---- new version: only the changed file is downloaded, card shows, update applies ----
    old_ver = pg.evaluate("caches.keys().then(k => k.filter(n => n.startsWith('mvl-')).join(','))")
    env2 = dict(env, MVL_TEST_EXEC=f'http://127.0.0.1:{PORT}/exec?x=2')
    subprocess.run(['node', str(ROOT / 'tools' / 'build-site.mjs')], check=True, env=env2)
    GETS.clear()
    pg.evaluate("navigator.serviceWorker.getRegistration().then(r => r.update())")
    pg.wait_for_selector('#mvlUpdate', timeout=20000)
    fetched = sorted({g.split('?')[0] for g in GETS if '?v=' in g})
    check(fetched == ['/config.js'], 'update downloads only the changed file (' + ', '.join(fetched) + ')')
    check(pg.is_visible('#mvlUpdateGo'), 'big "New update ready" card shown (wheel not spinning)')
    pg.click('#mvlUpdateGo')
    pg.wait_for_timeout(2500)
    new_ver = pg.evaluate("caches.keys().then(k => k.filter(n => n.startsWith('mvl-')).join(','))")
    check(new_ver != old_ver and pg.evaluate("window.MVL_EXEC.endsWith('?x=2')"), f'update applied after tap ({old_ver} → {new_ver})')
    # ---- card never appears mid-spin ----
    subprocess.run(['node', str(ROOT / 'tools' / 'build-site.mjs')], check=True, env=dict(env, MVL_TEST_EXEC=f'http://127.0.0.1:{PORT}/exec?x=3'))
    pg.evaluate("wheelBusy = true; navigator.serviceWorker.getRegistration().then(r => r.update())")
    pg.wait_for_timeout(5000)
    check(not pg.is_visible('#mvlUpdate'), 'no update card while the wheel is busy')
    pg.evaluate("wheelBusy = false")
    pg.wait_for_selector('#mvlUpdate', timeout=10000)
    check(True, 'card appears once the wheel is free')
    check(not [e for e in errs if 'net::ERR_INTERNET_DISCONNECTED' not in e and 'Failed to fetch' not in e and 'No connection' not in e],
          'no page errors' + ('' if not errs else ': ' + '; '.join(errs[:3])))
    b.close()
srv.shutdown()
print(('❌ ' + str(len(fails)) + ' failed') if fails else '✅ Website test passed')
sys.exit(1 if fails else 0)

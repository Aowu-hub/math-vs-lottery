/* Service worker: keeps the app's own files on the phone so it opens
   instantly; data still comes live from the app's server. A new version
   only downloads the files whose fingerprint changed. */
var VERSION = '__VERSION__';
var FILES = __FILES__;
var CACHE = 'mvl-' + VERSION;

self.addEventListener('install', function (e) {
  e.waitUntil((async function () {
    var cache = await caches.open(CACHE);
    var oldNames = (await caches.keys()).filter(function (n) { return n !== CACHE && n.indexOf('mvl-') === 0; });
    for (var path in FILES) {
      var reused = false;
      for (var i = 0; i < oldNames.length && !reused; i++) {
        var old = await caches.open(oldNames[i]);
        var meta = await old.match('__files__');
        if (!meta) continue;
        var map = await meta.json();
        if (map[path] === FILES[path]) {
          var hit = await old.match(path);
          if (hit) { await cache.put(path, hit); reused = true; }
        }
      }
      if (!reused) await cache.put(path, await fetch(path + '?v=' + FILES[path], { cache: 'no-store' }).then(function (r) { if (!r.ok) throw new Error(path); return r; }));
    }
    await cache.put('__files__', new Response(JSON.stringify(FILES), { headers: { 'Content-Type': 'application/json' } }));
  })());
});

self.addEventListener('activate', function (e) {
  e.waitUntil((async function () {
    var names = await caches.keys();
    await Promise.all(names.filter(function (n) { return n.indexOf('mvl-') === 0 && n !== CACHE; }).map(function (n) { return caches.delete(n); }));
    await self.clients.claim();
  })());
});

self.addEventListener('message', function (e) { if (e.data === 'skip') self.skipWaiting(); });

self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return; // the app's server (/exec) is never cached
  var scope = new URL(self.registration.scope);
  var path = url.pathname.slice(scope.pathname.length) || 'index.html';
  if (req.mode === 'navigate') path = 'index.html';
  e.respondWith(caches.open(CACHE).then(function (cache) {
    return cache.match(path).then(function (hit) { return hit || fetch(req); });
  }));
});

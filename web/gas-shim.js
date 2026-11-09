/* Lets the same app page run outside Apps Script (GitHub Pages website):
   google.script.run.xxx() becomes a POST to the app's /exec "website door"
   ({api:'run', fn, args}); google.script.url.getLocation() reads this URL. */
(function () {
  var EXEC = window.MVL_EXEC;
  function post(fn, args, tries) {
    return fetch(EXEC, {
      method: 'POST', redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ api: 'run', fn: fn, args: args })
    }).then(function (r) {
      if (!r.ok) throw new Error('HTTP ' + r.status);
      return r.json();
    }).catch(function (e) {
      if (tries > 0) return new Promise(function (res) { setTimeout(res, 1200); }).then(function () { return post(fn, args, tries - 1); });
      throw e;
    });
  }
  function runner() {
    var h = {};
    var p = new Proxy({}, {
      get: function (t, k) {
        if (k === 'withSuccessHandler') return function (f) { h.s = f; return p; };
        if (k === 'withFailureHandler') return function (f) { h.f = f; return p; };
        if (k === 'withUserObject') return function () { return p; };
        return function () {
          var args = Array.prototype.slice.call(arguments);
          post(String(k), args, 1).then(function (d) {
            if (d && d.ok) { if (h.s) h.s(d.result); }
            else { var err = new Error((d && d.error) || 'Server error'); if (h.f) h.f(err); else console.error(err); }
          }, function () {
            var err = new Error('📶 ต่ออินเทอร์เน็ตไม่ได้ ลองใหม่อีกครั้ง · No connection — check the internet and try again.');
            if (h.f) h.f(err); else console.error(err);
          });
        };
      }
    });
    return p;
  }
  window.google = {
    script: {
      get run() { return runner(); },
      url: {
        getLocation: function (cb) {
          var one = {}, all = {};
          new URLSearchParams(location.search).forEach(function (v, k) { if (!(k in one)) one[k] = v; (all[k] = all[k] || []).push(v); });
          setTimeout(function () { cb({ parameter: one, parameters: all, hash: location.hash.replace(/^#/, '') }); }, 0);
        }
      },
      host: { close: function () {}, setHeight: function () {}, setWidth: function () {} }
    }
  };
})();

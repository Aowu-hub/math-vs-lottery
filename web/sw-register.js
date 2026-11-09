/* Installable + fast: registers the service worker and shows a big
   "new update ready" card — only while the wheel is NOT spinning. */
(function () {
  if (!('serviceWorker' in navigator) || location.protocol === 'file:') return;
  var shown = false;
  function wheelBusyNow() { try { return typeof wheelBusy !== 'undefined' && wheelBusy; } catch (e) { return false; } }
  function showCard(reg) {
    if (shown) return;
    if (wheelBusyNow() || document.getElementById('welcomeOverlay')) { setTimeout(function () { showCard(reg); }, 2000); return; }
    shown = true;
    var box = document.createElement('div');
    box.id = 'mvlUpdate';
    box.setAttribute('role', 'dialog');
    box.style.cssText = 'position:fixed;inset:0;z-index:500;display:flex;align-items:center;justify-content:center;padding:24px;background:rgba(5,8,22,.72)';
    box.innerHTML = '<div style="max-width:360px;width:100%;background:#1f1a55;border:1px solid rgba(170,150,255,.35);border-radius:22px;padding:24px 20px;text-align:center;color:#efeaff;font-family:Sarabun,system-ui,sans-serif">' +
      '<div style="font-size:42px">🆕</div><div style="font:600 20px Mitr,Sarabun,sans-serif;margin:6px 0 2px">มีอัปเดตใหม่</div>' +
      '<div style="color:#b9b0e6;margin-bottom:16px">New update ready</div>' +
      '<button type="button" id="mvlUpdateGo" style="width:100%;min-height:50px;border:0;border-radius:14px;background:linear-gradient(135deg,#8f7bff,#f59fd6);color:#1b1240;font:700 17px Sarabun,sans-serif;cursor:pointer">⟳ อัปเดตเลย · Update now</button>' +
      '<button type="button" id="mvlUpdateLater" style="margin-top:10px;background:none;border:0;color:#b9b0e6;font:15px Sarabun,sans-serif;min-height:44px;cursor:pointer">ภายหลัง · Later</button></div>';
    document.body.appendChild(box);
    document.getElementById('mvlUpdateGo').onclick = function () { updating = true; if (reg.waiting) reg.waiting.postMessage('skip'); else location.reload(); };
    document.getElementById('mvlUpdateLater').onclick = function () { box.remove(); };
  }
  // Reload only when the person tapped "Update now" — not on the very first
  // visit, when the new service worker simply takes over the open page.
  var updating = false, reloading = false;
  navigator.serviceWorker.addEventListener('controllerchange', function () { if (updating && !reloading) { reloading = true; location.reload(); } });
  window.addEventListener('load', function () {
    navigator.serviceWorker.register('sw.js').then(function (reg) {
      function watch(w) { w.addEventListener('statechange', function () { if (w.state === 'installed' && navigator.serviceWorker.controller) showCard(reg); }); }
      if (reg.waiting && navigator.serviceWorker.controller) showCard(reg);
      if (reg.installing) watch(reg.installing);
      reg.addEventListener('updatefound', function () { if (reg.installing) watch(reg.installing); });
      setInterval(function () { reg.update().catch(function () {}); }, 30 * 60 * 1000);
      document.addEventListener('visibilitychange', function () { if (!document.hidden) reg.update().catch(function () {}); });
    }).catch(function () {});
  });
})();

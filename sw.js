importScripts('https://cdn.onesignal.com/sdks/web/v16/OneSignalSDKWorker.js');
/* Outlet Operation System - service worker (cache untuk penggunaan offline).
   Strategi: network-first untuk berkas aplikasi di domain yang sama, sehingga pengguna online
   selalu mendapat versi terbaru; cache hanya dipakai saat jaringan tidak tersedia.
   Permintaan API (Apps Script, lintas domain, POST) tidak pernah disentuh.
   Saat merilis versi baru, naikkan CACHE_VERSION. */
var CACHE_VERSION = 'nova-v20.3.2-repo-tidy';
var CACHE = CACHE_VERSION;
var SHELL = [
  './',
  './index.html',
  './assets/css/tailwind.css',
  './assets/css/legacy-modules.css',
  './assets/css/oos.css',
  './assets/css/figma.css',
  './assets/css/nova.css',
  './assets/css/nova-ds.css',
  './assets/css/nova-figma.css',
  './assets/js/multi-item.js',
  './assets/js/pdfkit.js',
  './assets/js/msel.js',
  './assets/img/logo/logo-light.png',
  './assets/img/logo/logo-xxi.png',
  './assets/img/logo/nova-logo.png',
  './assets/img/logo/nova-logo-light.png',
  './assets/js/ac-trend.js',
  './assets/js/loader.js',
  './assets/js/shell.js',
  './assets/js/app.js',
  './assets/js/sidebar.js',
  './assets/js/ux.js',
  './assets/js/drafts.js',
  './assets/js/mode.js',
  './assets/js/notify.js',
  './assets/js/nav-back.js',
  './assets/js/modul-dcp.js',
  './assets/js/pelaporan.js',
  './assets/js/penonton.js',
  './assets/img/cut-christie.webp',
  './assets/img/cut-barco.webp',
  './manifest.json',
  './assets/img/logo/logo.png',
  './assets/icons/icon-192.png',
  './assets/icons/icon-512.png',
  './assets/icons/icon-512-maskable.png',
  './assets/icons/apple-touch-icon.png'
];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (c) {
      return Promise.all(SHELL.map(function (u) {
        return c.add(u).catch(function () {});
      }));
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== CACHE; }).map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

/* network-first dengan batas waktu: jika jaringan lambat (> 3,5 detik) dan ada salinan cache, tampilkan cache lebih dulu
   sementara permintaan jaringan tetap berjalan dan memperbarui cache untuk kunjungan berikutnya. Permintaan API tidak pernah disentuh. */
var NET_TIMEOUT_MS = 3500;
self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET') return;
  var url = new URL(req.url);
  if (url.origin !== self.location.origin) return; /* API & CDN: biarkan langsung ke jaringan */

  var fromNet = fetch(req).then(function (res) {
    if (res && res.status === 200 && res.type === 'basic') {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(req, copy); });
    }
    return res;
  });

  var fallback = function () {
    return caches.match(req, { ignoreSearch: true }).then(function (hit) {
      if (hit) return hit;
      if (req.mode === 'navigate') return caches.match('./index.html');
      return Response.error();
    });
  };

  var slow = new Promise(function (resolve) {
    setTimeout(function () {
      caches.match(req, { ignoreSearch: true }).then(function (hit) { if (hit) resolve(hit); });
    }, NET_TIMEOUT_MS);
  });

  e.respondWith(Promise.race([fromNet.catch(fallback), slow]));
});

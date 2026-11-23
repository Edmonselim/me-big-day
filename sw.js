// Offline support: keeps the app itself on the phone. Workbook data lives in IndexedDB (see store.js).
const VERSION = 'meb-v1';
const SHELL = [
  './', './index.html', './config.js', './manifest.webmanifest', './css/app.css',
  './js/app.js', './js/ui.js', './js/store.js', './js/model.js', './js/xlsx.js', './js/zip.js', './js/graph.js', './js/seating.js',
  './img/couple.jpg', './img/icon-192.png', './img/icon-512.png', './img/apple-touch-icon.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

// Same-origin files: network first (so updates arrive), fall back to cache when offline.
// Everything else (Microsoft sign-in, Graph, exchange rates) goes straight to the network.
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== location.origin) return;
  if (url.search.includes('code=')) return; // sign-in redirect
  e.respondWith(
    fetch(e.request).then((res) => {
      if (res.ok) { const copy = res.clone(); caches.open(VERSION).then((c) => c.put(e.request, copy)); }
      return res;
    }).catch(() => caches.match(e.request, { ignoreSearch: true }).then((r) => r || caches.match('./index.html'))),
  );
});

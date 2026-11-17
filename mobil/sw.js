// Offline gyorsítótár az app fájljaihoz; az API hívásokat nem érinti.
const CACHE = 'arr-jelentes-v3.3';
const FILES = ['./', 'index.html', 'app.js', 'pdfparse.js', 'xlsm.js', 'manifest.webmanifest',
  'vendor/jszip.min.js', 'vendor/pdf.min.mjs', 'vendor/pdf.worker.min.mjs',
  'icons/icon-180.png', 'icons/icon-192.png', 'icons/icon-512.png', 'sablon.xlsm', 'utmutato.html'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (url.origin !== location.origin || e.request.method !== 'GET') return;
  e.respondWith(fetch(e.request).then(r => {
    const copy = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); return r;
  }).catch(() => caches.match(e.request, { ignoreSearch: true })));
});

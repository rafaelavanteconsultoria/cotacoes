const C = 'cotacoes-v1';
const FILES = ['./', 'index.html', 'manifest.webmanifest', 'lib/jspdf.umd.min.js', 'header.jpg',
  'icons/icon-192.png', 'icons/icon-512.png', 'icons/icon-180.png', 'icons/icon-maskable-512.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(C).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== C).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request).then(m => {
    const net = fetch(e.request).then(r => {
      if (r && r.ok && new URL(e.request.url).origin === location.origin) {
        const cp = r.clone();
        caches.open(C).then(c => c.put(e.request, cp));
      }
      return r;
    }).catch(() => m);
    return m || net;
  }));
});

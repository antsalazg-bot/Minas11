// ════════════════════════════════════════════════════════════════
// SERVICE WORKER — Caseta Minas 11
// Desarrollado por Antonio Salazar
// ════════════════════════════════════════════════════════════════
const CACHE = 'caseta-v1';
const ASSETS = [
  '/Minas11/caseta.html',
  '/Minas11/manifest-caseta.json',
  '/Minas11/icons/icon-caseta-192.png',
  '/Minas11/icons/icon-caseta-512.png',
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // No cachear peticiones POST (llamadas al GAS)
  if (e.request.method !== 'GET') return;
  // No cachear peticiones externas (CDN, GAS, etc.)
  if (!e.request.url.startsWith(self.location.origin)) return;

  e.respondWith(
    caches.match(e.request).then(cached => {
      const fetchFresh = fetch(e.request).then(res => {
        if (res && res.status === 200) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => cached);
      return cached || fetchFresh;
    })
  );
});

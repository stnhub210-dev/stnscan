/* STN 스캔관리 PWA 서비스워커 — 네트워크 우선(항상 최신), 오프라인 시 캐시 폴백 */
const CACHE = 'stnscan-v1';
const SHELL = ['/index.html', '/icon-192.png', '/icon-512.png', '/manifest.json'];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  e.respondWith(
    fetch(e.request)
      .then((r) => {
        const cp = r.clone();
        caches.open(CACHE).then((c) => c.put(e.request, cp));
        return r;
      })
      .catch(() => caches.match(e.request).then((m) => m || caches.match('/index.html')))
  );
});

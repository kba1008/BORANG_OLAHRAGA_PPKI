/**
 * Service Worker - Sistem Sijil Pro v3
 * Strategi:
 *  - App shell (HTML/manifest/ikon): cache-first + kemas kini di latar (stale-while-revalidate)
 *  - Library CDN: cache-first (laju & boleh offline)
 *  - Panggilan API Apps Script: SENTIASA rangkaian (jangan cache data pelajar)
 */
const CACHE_NAME = 'sijilpro-v25';
const SHELL = ['./', './index.html', './manifest.json'];
const CDN_HOSTS = ['cdnjs.cloudflare.com', 'docs.opencv.org', 'fonts.googleapis.com', 'fonts.gstatic.com'];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return; // POST ke Apps Script - biar terus ke rangkaian

  const url = new URL(req.url);

  // Jangan cache API / thumbnail Drive (data sentiasa terkini)
  if (url.hostname.indexOf('script.google.com') > -1 || url.hostname.indexOf('drive.google.com') > -1) return;

  if (CDN_HOSTS.indexOf(url.hostname) > -1) {
    event.respondWith(
      caches.match(req).then(hit => hit || fetch(req).then(res => {
        const copy = res.clone();
        caches.open(CACHE_NAME).then(c => c.put(req, copy));
        return res;
      }))
    );
    return;
  }

  // App shell: stale-while-revalidate
  event.respondWith(
    caches.match(req).then(hit => {
      const network = fetch(req).then(res => {
        if (res && res.status === 200 && res.type === 'basic') {
          const copy = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(req, copy));
        }
        return res;
      }).catch(() => hit);
      return hit || network;
    })
  );
});

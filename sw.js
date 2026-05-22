const CACHE = 'htm-shipping-v1';
const PRECACHE = ['./', './index.html', './icon.svg'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(PRECACHE)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const url = new URL(e.request.url);
  // Never intercept Firebase/Firestore — those need live network
  if (url.hostname.includes('googleapis.com') && url.pathname.includes('firestore')) return;
  if (url.hostname.includes('firebase') && !url.hostname.includes('fonts')) return;

  e.respondWith(
    caches.match(e.request).then(hit => {
      if (hit) return hit;
      return fetch(e.request).then(res => {
        // Cache fonts and Firebase SDK scripts for offline
        if (
          url.hostname.includes('fonts.gstatic.com') ||
          url.hostname.includes('fonts.googleapis.com') ||
          (url.hostname.includes('gstatic.com') && url.pathname.includes('firebasejs'))
        ) {
          const clone = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, clone));
        }
        return res;
      }).catch(() => {
        if (e.request.mode === 'navigate') return caches.match('./index.html');
      });
    })
  );
});

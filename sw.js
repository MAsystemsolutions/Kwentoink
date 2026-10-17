/* Kwentoink offline helper.
   Keeps the app saved on the phone so it opens instantly and works without internet.
   When you upload a new index.html, change the version number below (v1 → v2) so phones update. */
const CACHE = 'kwentoink-v2';
const FILES = ['./', './index.html', './manifest.json', './logo.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);
  // Only handle our own files. Data calls to Google go straight to the internet.
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;

  // Show the saved copy right away, refresh it in the background.
  e.respondWith(
    caches.open(CACHE).then(cache =>
      cache.match(req, { ignoreSearch: true }).then(saved => {
        const fresh = fetch(req).then(res => {
          if (res && res.ok) cache.put(req.mode === 'navigate' ? './index.html' : req, res.clone());
          return res;
        }).catch(() => saved || cache.match('./index.html'));
        return saved || fresh;
      })
    )
  );
});

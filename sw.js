/* Kwentoink offline helper.
   Keeps the app saved on the phone so it opens instantly and works without internet.
   When you upload a new index.html, change the version number below (v5 → v6) so phones update. */
const CACHE = 'kwentoink-v5';
const FILES = ['./', './index.html', './manifest.json', './logo.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES.map(f => new Request(f, { cache: 'reload' })))).then(() => self.skipWaiting()));
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

  // The app page: get the newest version when online (so updates show right away),
  // fall back to the saved copy when offline or the connection is very slow.
  if (req.mode === 'navigate') {
    e.respondWith((async () => {
      const cache = await caches.open(CACHE);
      try {
        const res = await Promise.race([
          fetch(req, { cache: 'no-store' }),
          new Promise((_, rej) => setTimeout(() => rej(new Error('slow')), 4000))
        ]);
        if (res && res.ok) { cache.put('./', res.clone()); cache.put('./index.html', res.clone()); }
        return res;
      } catch (err) {
        return (await cache.match('./')) || (await cache.match('./index.html')) || Response.error();
      }
    })());
    return;
  }

  // Logo, manifest: show the saved copy right away, refresh it in the background.
  e.respondWith(
    caches.open(CACHE).then(cache =>
      cache.match(req, { ignoreSearch: true }).then(saved => {
        const fresh = fetch(req).then(res => {
          if (res && res.ok) cache.put(req, res.clone());
          return res;
        }).catch(() => saved);
        return saved || fresh;
      })
    )
  );
});

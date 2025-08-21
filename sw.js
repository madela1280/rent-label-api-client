self.addEventListener('install', (e) => {
  e.waitUntil((async () => {
    const cache = await caches.open('app-v1');
    await cache.addAll(['./', './index.html', './manifest.webmanifest']);
  })());
});
self.addEventListener('fetch', (e) => {
  e.respondWith((async () => {
    const cached = await caches.match(e.request);
    return cached || fetch(e.request);
  })());
});

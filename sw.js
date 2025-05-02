const CACHE_NAME = 'todo-pwa-v4'; // Update version

// Install event: Cache essential files and activate immediately
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll([
        '/todo-pwa/',
        '/todo-pwa/index.html',
        '/todo-pwa/manifest.json',
        '/todo-pwa/icon-192x192.png',
        '/todo-pwa/icon-512x512.png'
      ]);
    })
  );
  self.skipWaiting(); // Activate new service worker immediately
});

// Activate event: Clear old caches and take control of clients
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.filter(name => name !== CACHE_NAME).map(name => caches.delete(name))
      );
    }).then(() => self.clients.claim()) // Take control of all open clients
  );
});

// Fetch event: Serve from cache, fallback to network
self.addEventListener('fetch', event => {
  event.respondWith(
    caches.match(event.request).then(response => {
      return response || fetch(event.request);
    })
  );
});

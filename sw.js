const CACHE_NAME = 'sparrow-farms-v8';
const urlsToCache = [
  '/',
  '/index.html'
];

// Install service worker and cache resources
self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(function(cache) {
        return cache.addAll(urlsToCache);
      })
  );
  self.skipWaiting(); // Activate immediately
});

// Serve cached content when offline - but NEVER cache API calls
self.addEventListener('fetch', function(event) {
  // Don't cache Google Apps Script API calls
  if (event.request.url.includes('script.google.com')) {
    return; // Let it fetch normally, don't cache
  }

  event.respondWith(
    fetch(event.request)
      .then(function(response) {
        // Update cache with fresh response
        const responseClone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, responseClone));
        return response;
      })
      .catch(function() {
        // Only fall back to cache when offline
        return caches.match(event.request);
      })
  );
});

// Clean up old caches
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(cacheNames) {
      return Promise.all(
        cacheNames.filter(function(cacheName) {
          return cacheName !== CACHE_NAME;
        }).map(function(cacheName) {
          return caches.delete(cacheName);
        })
      );
    })
  );
  self.clients.claim(); // Take control immediately
});

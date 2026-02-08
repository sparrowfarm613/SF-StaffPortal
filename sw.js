const CACHE_NAME = 'sparrow-farms-v3';
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
});

// Serve cached content when offline - but NEVER cache API calls
self.addEventListener('fetch', function(event) {
  // Don't cache Google Apps Script API calls
  if (event.request.url.includes('script.google.com')) {
    return; // Let it fetch normally, don't cache
  }
  
  event.respondWith(
    caches.match(event.request)
      .then(function(response) {
        return response || fetch(event.request);
      }
    )
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
});

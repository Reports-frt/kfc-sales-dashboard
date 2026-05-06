// =====================================================================
// KFC Dashboard Service Worker
// =====================================================================
// Strategy:
//   - HTML/CSS/JS (app shell): cache-first with background update
//   - data.json: network-first with fallback to cache (always try fresh)
//   - Icons & manifest: cache-first
// =====================================================================

const CACHE_VERSION = 'kfc-dashboard-v1';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-180.png',
  './favicon.png',
];

// On install: cache the app shell
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then(cache => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

// On activate: clean up old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// On fetch: serve appropriately based on resource type
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  
  // Only handle same-origin requests
  if (url.origin !== self.location.origin) return;

  const isDataJson = url.pathname.endsWith('/data.json');
  
  if (isDataJson) {
    // Network-first for data.json (always try to get fresh data)
    event.respondWith(
      fetch(event.request)
        .then(response => {
          // Clone and cache the fresh response
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_VERSION).then(cache => cache.put(event.request, clone));
          }
          return response;
        })
        .catch(() => caches.match(event.request))  // fallback to cached version
    );
  } else {
    // Cache-first for app shell (HTML, JS, CSS, icons)
    event.respondWith(
      caches.match(event.request)
        .then(cached => {
          if (cached) {
            // Background update: fetch fresh and update cache
            fetch(event.request)
              .then(fresh => {
                if (fresh.ok) {
                  caches.open(CACHE_VERSION).then(cache => cache.put(event.request, fresh));
                }
              })
              .catch(() => {});
            return cached;
          }
          // Not in cache, fetch normally
          return fetch(event.request);
        })
    );
  }
});

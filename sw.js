// BEE Task System — service worker
// Purpose: make the app shell load instantly / installable.
// Does NOT cache anything from script.google.com — task data is always
// fetched live, never served stale from cache.

const CACHE_NAME = 'bee-tasks-shell-v1';
const SHELL_FILES = [
  './index.html',
  './bbbee_v5_2.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-192-maskable.png',
  './icon-512-maskable.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(
        names
          .filter((name) => name !== CACHE_NAME)
          .map((name) => caches.delete(name))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = event.request.url;

  // Never intercept calls to the Apps Script backend (task data, JSONP).
  // Always go straight to the network so the board is always live.
  if (url.includes('script.google.com') || url.includes('script.googleusercontent.com')) {
    return; // let the browser handle it normally, no caching
  }

  // For everything else (the app shell + CDN fonts/Chart.js):
  // try the network first so updates are picked up quickly,
  // fall back to cache if offline.
  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // Only cache same-origin shell files, not opaque CDN errors etc.
        if (event.request.method === 'GET' && response && response.status === 200) {
          const responseClone = response.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseClone);
          });
        }
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});

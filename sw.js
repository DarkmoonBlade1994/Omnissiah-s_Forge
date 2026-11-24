/**
 * Home Projects & Groceries — Service Worker
 * Caches the app shell for instant mobile launch & offline shell availability,
 * while always going directly to the network for Google Apps Script API data
 * so multiple users stay in sync against the shared Google Sheet.
 */

const CACHE_NAME = 'home-dashboard-shell-v4';
const SHELL_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon.svg',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => key !== CACHE_NAME)
          .map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // Never cache Google Apps Script API requests — always fetch fresh shared Sheet data
  if (
    url.hostname.endsWith('script.google.com') ||
    url.hostname.endsWith('script.googleusercontent.com')
  ) {
    event.respondWith(fetch(req));
    return;
  }

  // Only cache same-origin app shell assets
  if (url.origin !== self.location.origin) return;

  // Network-first for HTML documents (keeps app shell fresh while falling back to cache offline)
  if (req.mode === 'navigate' || req.destination === 'document') {
    event.respondWith(
      fetch(req)
        .then((resp) => {
          if (resp && resp.ok) {
            const copy = resp.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
          }
          return resp;
        })
        .catch(() => caches.match(req).then((cached) => cached || caches.match('./index.html')))
    );
    return;
  }

  // Cache-first with background refresh for icons and manifest
  event.respondWith(
    caches.match(req).then((cached) => {
      const networkFetch = fetch(req)
        .then((resp) => {
          if (resp && resp.ok) {
            const copy = resp.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(req, copy));
          }
          return resp;
        })
        .catch(() => cached);

      return cached || networkFetch;
    })
  );
});

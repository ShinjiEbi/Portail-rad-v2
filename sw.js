// ══════════════════════════════════════════════════════════
// Service Worker – Portail Tech Bertin
// Version auto-incrémentée par export_data.py
// ══════════════════════════════════════════════════════════
const VERSION = 'fb3c78527fa1';
const CACHE   = 'portail-tech-' + VERSION;

// Fichiers à cacher (mis à jour automatiquement par export_data.py)
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './forms/bcp31_sonde_beta.html',
];

// ── INSTALL : pré-cache tous les assets ─────────────────
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(cache) {
      return Promise.allSettled(
        ASSETS.map(function(url) {
          return cache.add(url).catch(function() {
            console.warn('[SW] Cache miss:', url);
          });
        })
      );
    })
  );
  self.skipWaiting();
});

// ── ACTIVATE : nettoyer les anciens caches ──────────────
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE; })
            .map(function(k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// ── MESSAGE : forcer l'activation depuis l'app ──────────
self.addEventListener('message', function(e) {
  if (e.data && e.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

// ── FETCH : cache-first, fallback network ───────────────
self.addEventListener('fetch', function(e) {
  // Ne pas intercepter les requêtes non-GET
  if (e.request.method !== 'GET') return;

  e.respondWith(
    caches.match(e.request).then(function(cached) {
      if (cached) return cached;
      return fetch(e.request).then(function(response) {
        if (!response || response.status !== 200) return response;
        var clone = response.clone();
        caches.open(CACHE).then(function(cache) {
          cache.put(e.request, clone);
        });
        return response;
      }).catch(function() {
        // Offline fallback : retourner index.html pour navigation
        return caches.match('./index.html');
      });
    })
  );
});

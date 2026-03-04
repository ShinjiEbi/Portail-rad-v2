// ══════════════════════════════════════════════════════════
// Service Worker – Portail Tech Bertin
// Stratégie : network-first pour index.html (détecte les MAJ)
//             cache-first pour le reste (offline)
// ══════════════════════════════════════════════════════════
const CACHE = 'portail-tech-v1';

// ── INSTALL ─────────────────────────────────────────────
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE).then(function(cache) {
      return cache.addAll(['./', './index.html', './manifest.json']).catch(function() {});
    })
  );
  // Ne PAS skipWaiting ici — on laisse l'app décider via SKIP_WAITING
});

// ── ACTIVATE : nettoyer les anciens caches ──────────────
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k.startsWith('portail-tech-') && k !== CACHE; })
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

// ── FETCH ───────────────────────────────────────────────
self.addEventListener('fetch', function(e) {
  if (e.request.method !== 'GET') return;

  var url = new URL(e.request.url);

  // Network-first pour index.html (détecte les MAJ à chaque ouverture)
  if (url.pathname.endsWith('/') || url.pathname.endsWith('index.html')) {
    e.respondWith(
      fetch(e.request).then(function(response) {
        if (response && response.status === 200) {
          var clone = response.clone();
          caches.open(CACHE).then(function(cache) { cache.put(e.request, clone); });
        }
        return response;
      }).catch(function() {
        return caches.match(e.request);
      })
    );
    return;
  }

  // Cache-first pour tout le reste (forms/, manifest, images...)
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      if (cached) return cached;
      return fetch(e.request).then(function(response) {
        if (!response || response.status !== 200) return response;
        var clone = response.clone();
        caches.open(CACHE).then(function(cache) { cache.put(e.request, clone); });
        return response;
      }).catch(function() {
        return caches.match('./index.html');
      });
    })
  );
});

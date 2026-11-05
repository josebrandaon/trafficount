// TraffiCount — guarda a aplicação no telemóvel para funcionar sem internet.
// Ao publicar uma nova versão, altere VERSION para os telemóveis receberem a atualização.
const VERSION = 'trafficount-v15';
const SHELL = [
  './', './index.html', './manifest.webmanifest',
  './vendor/leaflet/leaflet.css', './vendor/leaflet/leaflet.js',
  './vendor/leaflet/images/marker-icon.png', './vendor/leaflet/images/marker-icon-2x.png',
  './vendor/leaflet/images/marker-shadow.png', './vendor/leaflet/images/layers.png', './vendor/leaflet/images/layers-2x.png',
  './vendor/inter/inter-latin-wght-normal.woff2',
  './icons/icon-192.png', './icons/icon-512.png', './icons/icon-maskable-512.png', './icons/favicon-32.png', './icons/apple-touch-icon.png'
];
const TILES = 'trafficount-mapa';
const TILE_MAX = 400;

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== VERSION && k !== TILES).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin === location.origin && (req.mode === 'navigate' || /\/(index\.html)?$|\.webmanifest$|\.json$/.test(url.pathname))) {
    // página principal: rede primeiro (versão mais recente); cache só sem internet
    e.respondWith(caches.open(VERSION).then(c => fetch(req, {cache:'no-store'}).then(r => {
      if (r.ok) c.put(req, r.clone()); return r;
    }).catch(() => c.match(req, {ignoreSearch:true}).then(hit => hit || c.match('./index.html')))));
    return;
  }
  if (url.origin === location.origin) {
    // restantes ficheiros: responde da cache e atualiza em segundo plano
    e.respondWith(caches.open(VERSION).then(c => c.match(req, {ignoreSearch:true}).then(hit => {
      const net = fetch(req).then(r => { if (r.ok) c.put(req, r.clone()); return r; }).catch(() => hit);
      return hit || net;
    })));
    return;
  }
  if (/tile\.openstreetmap\.org|arcgisonline\.com/.test(url.hostname)) {
    // mapa: rede primeiro; guarda as últimas zonas vistas para uso sem internet
    e.respondWith(fetch(req).then(r => {
      if (r.ok || r.type === 'opaque') caches.open(TILES).then(c => c.put(req, r.clone()).then(() => c.keys()).then(k => { if (k.length > TILE_MAX) return c.delete(k[0]); }));
      return r;
    }).catch(() => caches.match(req)));
  }
});

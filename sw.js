const CACHE = 'umeapp-v3';
const CORE = [
  './', './index.html', './mapa.html', './bomberos.html', './cartelera.html', './gas.html',
  './manifest.json', './icon-192.png', './icon-512.png',
  './assets/centro-umepay.jpg'
];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;            // no tocar los envíos (pedidos de gas, etc.)
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;  // CDN, mapas y hojas van directo a la red
  e.respondWith(
    fetch(req).then(res => {
      const copia = res.clone();
      caches.open(CACHE).then(c => c.put(req, copia));
      return res;
    }).catch(() => caches.match(req))
  );
});

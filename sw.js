// Cache para funcionar sin conexión. Subir la versión al cambiar archivos.
const CACHE = 'lotes-villalobos-v7';
const ARCHIVOS = ['./', 'index.html', 'config.js', 'lotes.js', 'lotes-magallanes.js', 'lotes-gaboto.js', 'mapa.jpg', 'mapa-magallanes.jpg', 'mapa-gaboto.jpg', 'manifest.webmanifest',
  'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];

self.addEventListener('install', e => {
  // cache: 'reload' evita guardar copias viejas del caché HTTP del navegador
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARCHIVOS.map(u => new Request(u, { cache: 'reload' }))))
    .then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
// Red primero (para recibir actualizaciones), cache si no hay conexión
self.addEventListener('fetch', e => {
  // Los propietarios (Apps Script) los maneja la app; acá solo archivos propios
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  e.respondWith(
    fetch(e.request, { cache: 'no-cache' }) // siempre revalida con el servidor
      .then(r => { const copia = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copia)); return r; })
      .catch(() => caches.match(e.request, { ignoreSearch: true }))
  );
});

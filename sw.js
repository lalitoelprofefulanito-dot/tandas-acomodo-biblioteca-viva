// Tandas de acomodo · Biblioteca Viva — funcionamiento sin internet.
// Guarda la página y cada portada que se va mostrando.
const CACHE = 'tandas-bv-2026-11-13-1794-v4';

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(['./', './index.html'])));
});

self.addEventListener('activate', e => e.waitUntil(
  caches.keys()
    .then(claves => Promise.all(claves.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim())
));

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  const guardar = r => {
    if (r.ok) { const copia = r.clone(); caches.open(CACHE).then(c => c.put(e.request, copia)); }
    return r;
  };
  // La página: primero la red (para recibir actualizaciones); sin internet, la copia guardada.
  // Las portadas: primero la copia guardada; si no existe, se descarga y se guarda.
  e.respondWith(e.request.mode === 'navigate'
    ? fetch(e.request).then(guardar).catch(() => caches.match(e.request).then(r => r || caches.match('./index.html')))
    : caches.match(e.request).then(r => r || fetch(e.request).then(guardar)));
});

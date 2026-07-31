// Service worker mínimo: no cachea nada, solo deja pasar las peticiones.
// Su única función es cumplir el requisito técnico de Android/Chrome para
// poder mostrar el aviso de "Instalar app".
self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  event.respondWith(fetch(event.request));
});

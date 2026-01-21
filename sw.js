/* Service Worker — Planejador de Recursos
 * Estratégia:
 * - App shell (arquivos essenciais) em cache na instalação
 * - Cache-first para navegação/app shell
 * - Stale-while-revalidate para demais assets
 */

const CACHE_VERSION = 'pr-v1';
const APP_SHELL = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './planner_enhancer.js',
  './enhancer.js',
  './enhancer2.js',
  './enhancer.css',
  './capacidade_v4.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/maskable-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.map((k) => (k !== CACHE_VERSION ? caches.delete(k) : Promise.resolve())))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Só controla mesmo origin
  if (url.origin !== self.location.origin) return;

  // Navegação: sempre devolver index.html (app shell) em modo cache-first
  if (req.mode === 'navigate') {
    event.respondWith(
      caches.match('./index.html').then((cached) => cached || fetch(req))
    );
    return;
  }

  // Para assets: stale-while-revalidate
  event.respondWith(
    caches.match(req).then((cached) => {
      const fetchPromise = fetch(req).then((resp) => {
        // Cache só respostas OK (evita cachear erros)
        if (resp && resp.ok) {
          const copy = resp.clone();
          caches.open(CACHE_VERSION).then((cache) => cache.put(req, copy));
        }
        return resp;
      }).catch(() => cached);

      return cached || fetchPromise;
    })
  );
});

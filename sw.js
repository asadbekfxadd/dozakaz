// Версия кэша — меняй при каждом деплое
const CACHE_VERSION = 'lining-v3';

// При установке - не кэшируем ничего
self.addEventListener('install', e => {
  self.skipWaiting();
});

// При активации - удаляем старые кэши
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.map(key => caches.delete(key)))
    ).then(() => self.clients.claim())
  );
});

// Все запросы идут в сеть, без кэша
self.addEventListener('fetch', e => {
  // Только GET запросы
  if (e.request.method !== 'GET') return;
  // API запросы всегда через сеть
  if (e.request.url.includes('/api/')) return;
  // Всё остальное тоже через сеть (без кэша)
  e.respondWith(fetch(e.request));
});

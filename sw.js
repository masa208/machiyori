const CACHE_NAME = 'machiyori-v2';
const ASSETS = [
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

// インストール時にキャッシュ
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// 古いキャッシュを削除
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const req = e.request;
  const url = new URL(req.url);

  // 自分のサイト以外（GAS・LINEなど）は素通し
  if (url.origin !== self.location.origin) return;

  const isHtml = req.mode === 'navigate' || url.pathname.endsWith('.html');

  if (isHtml) {
    // HTMLはネットワーク優先（更新を確実に反映）
    e.respondWith(
      fetch(req)
        .then(res => {
          const copy = res.clone();
          caches.open(CACHE_NAME).then(c => c.put(req, copy));
          return res;
        })
        .catch(() => caches.match(req))
    );
    return;
  }

  // それ以外はキャッシュ優先
  e.respondWith(
    caches.match(req).then(cached => cached || fetch(req))
  );
});

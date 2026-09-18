/* impactTV 端末管理 — Service Worker
 * GET のみ対象。API（POST）は素通し。
 * ナビゲーション: network-first / CDN: stale-while-revalidate / 同一オリジン: cache-first */
const VERSION = 'itv-v2.0.0';
const PRECACHE = [
  './',
  'index.html',
  'manifest.webmanifest',
  'config.js',
  'demo.js',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/maskable-512.png',
  'icons/apple-touch-icon.png',
  'icons/favicon-32.png'
];
const CDN_HOSTS = ['fastly.jsdelivr.net', 'cdn.jsdelivr.net', 'unpkg.com'];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(VERSION).then(function (cache) {
      // 1つ欠けてもインストール自体は成功させる
      return Promise.allSettled(PRECACHE.map(function (url) { return cache.add(url); }));
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (k) { return k !== VERSION; }).map(function (k) { return caches.delete(k); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (event) {
  const req = event.request;
  if (req.method !== 'GET') return; // API（POST）は素通し

  let url;
  try { url = new URL(req.url); } catch (e) { return; }

  // 画面遷移: network-first（だめならキャッシュの index.html）
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req).then(function (res) {
        const copy = res.clone();
        caches.open(VERSION).then(function (c) { c.put('index.html', copy); });
        return res;
      }).catch(function () {
        return caches.match('index.html').then(function (r) { return r || caches.match('./'); });
      })
    );
    return;
  }

  // CDN の読み取りライブラリ: stale-while-revalidate
  if (CDN_HOSTS.indexOf(url.host) >= 0) {
    event.respondWith(
      caches.match(req).then(function (cached) {
        const fetched = fetch(req).then(function (res) {
          if (res && (res.ok || res.type === 'opaque')) {
            const copy = res.clone();
            caches.open(VERSION).then(function (c) { c.put(req, copy); });
          }
          return res;
        }).catch(function () { return cached; });
        return cached || fetched;
      })
    );
    return;
  }

  // 同一オリジンのその他 GET: cache-first（なければ network → キャッシュ保存）
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(req).then(function (cached) {
        if (cached) return cached;
        return fetch(req).then(function (res) {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(VERSION).then(function (c) { c.put(req, copy); });
          }
          return res;
        });
      })
    );
  }
});

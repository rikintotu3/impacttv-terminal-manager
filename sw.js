/* impactTV 端末管理 — Service Worker
 * GET のみ対象。API（POST）は素通し。
 * 設定と画面: network-first / CDN: stale-while-revalidate / 静的資産: cache-first */
const VERSION = 'itv-v2.1.0';
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
const CONFIG_URL = new URL('config.js', self.location.href).href;
const STATIC_URLS = new Set(PRECACHE.map(function (path) {
  return new URL(path, self.location.href).href;
}));
const PAGE_PATHS = ['./', 'index.html'].map(function (path) {
  return new URL(path, self.location.href).pathname;
});
const CDN_HOSTS = ['fastly.jsdelivr.net', 'cdn.jsdelivr.net', 'unpkg.com'];

function isOwnCache(name) {
  return name.indexOf('itv-') === 0;
}

async function saveResponse(key, response) {
  // キャッシュ容量不足でも、取得できたレスポンスは利用できる。
  try {
    const cache = await caches.open(VERSION);
    await cache.put(key, response.clone());
  } catch (error) { /* キャッシュ保存はベストエフォート */ }
  return response;
}

async function fetchConfig(request) {
  const cache = await caches.open(VERSION);
  let response;
  let networkError;
  try {
    // HTTP キャッシュにも残さず、API 接続先の変更を毎回確認する。
    response = await fetch(request, { cache: 'no-store' });
    if (response.ok) return saveResponse(CONFIG_URL, response);
  } catch (error) {
    networkError = error;
  }
  const cached = await cache.match(CONFIG_URL);
  if (cached && cached.ok) return cached;
  // 設定がない場合、空の設定を成功扱いにしてデモへ切り替えない。
  if (response) return response;
  throw networkError;
}

async function preservePreviousConfig(cache) {
  if (await cache.match(CONFIG_URL)) return;
  const keys = (await caches.keys()).filter(function (name) {
    return isOwnCache(name) && name !== VERSION;
  }).reverse();
  for (const name of keys) {
    const previous = await caches.open(name);
    const config = await previous.match(CONFIG_URL);
    if (config && config.ok) {
      await cache.put(CONFIG_URL, config);
      return;
    }
  }
}

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(VERSION).then(async function (cache) {
      // 更新時に通信が失敗しても、最後に使えた設定を保持する。
      await preservePreviousConfig(cache);
      // 1つ欠けてもインストール自体は成功させる。
      return Promise.allSettled(PRECACHE.map(function (path) {
        const url = new URL(path, self.location.href).href;
        return url === CONFIG_URL ? fetchConfig(new Request(url)) : cache.add(url);
      }));
    }).then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(keys.filter(function (name) {
        return isOwnCache(name) && name !== VERSION;
      }).map(function (name) { return caches.delete(name); }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (event) {
  const req = event.request;
  if (req.method !== 'GET') return; // API（POST）は素通し。

  let url;
  try { url = new URL(req.url); } catch (error) { return; }

  if (url.origin === self.location.origin && url.pathname === new URL(CONFIG_URL).pathname) {
    event.respondWith(fetchConfig(req));
    return;
  }

  // アプリの画面遷移: network-first（通信不可なら保存済み画面）。
  if (req.mode === 'navigate' && url.origin === self.location.origin && PAGE_PATHS.indexOf(url.pathname) >= 0) {
    event.respondWith(
      fetch(req).then(function (res) {
        return res.ok ? saveResponse('index.html', res) : res;
      }).catch(async function () {
        const cache = await caches.open(VERSION);
        const cached = await cache.match('index.html') || await cache.match('./');
        if (cached) return cached;
        throw new Error('Offline page is unavailable');
      })
    );
    return;
  }

  // CDN の読み取りライブラリ: stale-while-revalidate。
  if (CDN_HOSTS.indexOf(url.host) >= 0) {
    const fetched = fetch(req).then(function (res) {
      return res && (res.ok || res.type === 'opaque') ? saveResponse(req, res) : res;
    }).catch(function () { return undefined; });
    event.waitUntil(fetched.then(function () {}));
    event.respondWith(
      caches.open(VERSION).then(async function (cache) {
        const response = await cache.match(req) || await fetched;
        if (response) return response;
        throw new Error('Offline library is unavailable');
      })
    );
    return;
  }

  // 既知の静的資産のみ保存し、業務データなどの GET は素通し。
  if (STATIC_URLS.has(url.href)) {
    event.respondWith(
      caches.open(VERSION).then(async function (cache) {
        const cached = await cache.match(req);
        if (cached) return cached;
        const response = await fetch(req);
        return response.ok ? saveResponse(req, response) : response;
      })
    );
  }
});

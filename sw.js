const CACHE_NAME = 'cz-smart-v12';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  'https://cdn.tailwindcss.com',
  'https://unpkg.com/lucide@latest'
];

// 安裝時快取資源
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

// 啟動時清理舊快取
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.map(key => {
        if (key !== CACHE_NAME) return caches.delete(key);
      })
    ))
  );
});

// 攔截請求
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);
  
  // 針對主頁面使用 Network-First 策略，確保 UI 邏輯隨時更新
  if (url.pathname === '/' || url.pathname.endsWith('index.html')) {
    event.respondWith(
      fetch(event.request)
        .then(response => {
          const clonedRes = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clonedRes));
          return response;
        })
        .catch(() => caches.match(event.request))
    );
  } else {
    // 其他資源使用 Cache-First 策略
    event.respondWith(
      caches.match(event.request)
        .then(response => response || fetch(event.request))
    );
  }
});

// 監聽推播訊息 (預留)
self.addEventListener('push', event => {
  const data = event.data ? event.data.json() : { title: '新公告', body: '崇正國樂團有新的公告內容！' };
  const options = {
    body: data.body,
    icon: 'icon-192.png',
    badge: 'icon-192.png',
    vibrate: [100, 50, 100],
    data: { url: './index.html' }
  };
  event.waitUntil(
    self.registration.showNotification(data.title, options)
  );
});

// 點擊通知開合 App
self.addEventListener('notificationclick', event => {
  event.notification.close();
  event.waitUntil(
    clients.openWindow(event.notification.data.url)
  );
});

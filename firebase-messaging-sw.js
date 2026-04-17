importScripts('https://www.gstatic.com/firebasejs/10.13.2/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.13.2/firebase-messaging-compat.js');

const firebaseConfig = {
  apiKey: "AIzaSyAjjkrPrHBsoYuazfHlS4DwclksIw5SYYk",
  authDomain: "baiyang-co.firebaseapp.com",
  projectId: "baiyang-co",
  storageBucket: "baiyang-co.firebasestorage.app",
  messagingSenderId: "744052473577",
  appId: "1:744052473577:web:ee5b767dfde9438af67a78",
  measurementId: "G-8CF9E1E98E"
};

try {
    firebase.initializeApp(firebaseConfig);
    const messaging = firebase.messaging();

    // v48: 使用相容版背景處理器 (排除 manual push 衝突)
    messaging.setBackgroundMessageHandler(function(payload) {
        console.log('[SW] Firebase 背景訊息解析: ', payload);
        const title = payload.notification?.title || payload.data?.title || '國樂團公告';
        const options = {
            body: payload.notification?.body || payload.data?.body || '崇正國樂團有新公告，請進入查看',
            icon: 'https://fe314343.github.io/0301/icon-192.png',
            badge: 'https://fe314343.github.io/0301/icon-192.png',
            data: { url: 'https://fe314343.github.io/0301/index.html' },
            tag: 'cz-broadcast', // 覆寫舊通知防止洗版
            renotify: true
        };
        return self.registration.showNotification(title, options);
    });
} catch (e) {
    console.error("Firebase init failed in SW", e);
}

const CACHE_NAME = 'cz-smart-v48';

self.addEventListener('install', event => { self.skipWaiting(); });

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.map(key => { if (key !== CACHE_NAME) return caches.delete(key); })
    )).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  event.respondWith(fetch(event.request).catch(() => caches.match(event.request)));
});

// v48: 我們移除手動 push 監聽器，讓上面的 setBackgroundMessageHandler 獨家處理
// 這能解決某些系統中「監聽器搶食」導致通知不跳的問題

self.addEventListener('notificationclick', event => {
  event.notification.close();
  event.waitUntil(clients.matchAll({ type: 'window' }).then(clientsArr => {
    if (clientsArr.length > 0) {
      clientsArr[0].focus();
      return clientsArr[0].navigate('./index.html');
    }
    return clients.openWindow('./index.html');
  }));
});


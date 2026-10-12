importScripts('https://www.gstatic.com/firebasejs/10.14.1/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.14.1/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey: 'AIzaSyAYUmDUQ-ft3hjVU7WBXyFvrMt2sXXEmoA',
  authDomain: 'junkai-chat.firebaseapp.com',
  projectId: 'junkai-chat',
  storageBucket: 'junkai-chat.firebasestorage.app',
  messagingSenderId: '592199751504',
  appId: '1:592199751504:web:d5e05642dc61610d33fe55'
});

var messaging = firebase.messaging();

// アプリを閉じている・裏にある間に届いた通知を表示し、バッジを更新する
messaging.onBackgroundMessage(function (payload) {
  var d = payload.data || {};
  return self.registration.showNotification(d.title || '巡回チャット', {
    body: d.body || '',
    icon: 'icon-192.png',
    data: { url: './' }
  }).then(function () {
    return self.registration.getNotifications();
  }).then(function (list) {
    if (self.navigator && self.navigator.setAppBadge) {
      return self.navigator.setAppBadge(list.length);
    }
  }).catch(function () {});
});

self.addEventListener('notificationclick', function (event) {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (list) {
      for (var i = 0; i < list.length; i++) {
        if ('focus' in list[i]) return list[i].focus();
      }
      return clients.openWindow('./');
    })
  );
});

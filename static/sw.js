/* Cofle Manage — Service Worker (Web Push)
   Operatör telefonuna bildirim: launch/referans eklendiğinde sunucu push gönderir,
   burası bildirimi gösterir. Tıklayınca uygulama açılır/odaklanır. */
'use strict';

self.addEventListener('install', function (e) { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });

self.addEventListener('push', function (e) {
  var d = {};
  try { d = e.data ? e.data.json() : {}; }
  catch (err) { d = { title: 'Cofle Manage', body: e.data ? e.data.text() : '' }; }
  e.waitUntil(self.registration.showNotification(d.title || 'Cofle Manage', {
    body: d.body || '',
    icon: '/logo',
    badge: '/logo',
    data: { url: d.url || '/' },
    vibrate: [180, 80, 180]
  }));
});

self.addEventListener('notificationclick', function (e) {
  e.notification.close();
  var url = (e.notification.data && e.notification.data.url) || '/';
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (list) {
    for (var i = 0; i < list.length; i++) {
      if ('focus' in list[i]) return list[i].focus();
    }
    return self.clients.openWindow(url);
  }));
});

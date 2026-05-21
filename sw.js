const CACHE_NAME = 'sv-k9z4a';
const GAS_URL = '__GAS_WEB_APP_URL__'; // デプロイ後に置換

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(clients.claim()));

self.addEventListener('push', e => {
  e.waitUntil((async () => {
    let data = { active: false };
    try {
      const res = await fetch(`${GAS_URL}?action=latest&t=${Date.now()}`);
      data = await res.json();
    } catch (_) {}

    if (!data.active) return;

    await self.registration.showNotification('⚡ センサー発動', {
      body: `${data.bank}  風${data.wind}m/s  天運${data.tenun}`,
      data,
      requireInteraction: true,
      tag: 'sensor-alert',
    });
  })());
});

self.addEventListener('notificationclick', e => {
  e.notification.close();
  const data = e.notification.data || {};
  e.waitUntil((async () => {
    const all = await clients.matchAll({ type: 'window', includeUncontrolled: true });
    const target = all.find(c => c.url.includes(self.registration.scope));
    if (target) {
      await target.postMessage({ type: 'SENSOR_HIT', data });
      return target.focus();
    }
    return clients.openWindow(self.registration.scope);
  })());
});

const CACHE_NAME = 'sv-k9z4b';
let pendingData = null;
const GAS_URL = 'https://script.google.com/macros/s/AKfycbw0PlcC7z3ixoFx86bu0Dyj2jn7pNasRzLmpjZFRTawzGzXxxp2-JRZg51XjS2XgvPgrg/exec';

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
    pendingData = data;
    return clients.openWindow(self.registration.scope);
  })());
});

self.addEventListener('message', e => {
  if (e.data?.type === 'CLIENT_READY' && pendingData) {
    e.source.postMessage({ type: 'SENSOR_HIT', data: pendingData });
    pendingData = null;
  }
});

const CACHE_NAME = 'sv-k9z4e';
const GAS_URL = 'https://script.google.com/macros/s/AKfycbxlQb8Yzi_6FDPGPYa7yOR5M4IqDrjh_af5ihA6tRUPa4qD67MQx6XSnKfNoNLOj4jCww/exec';
const VAPID_PUBLIC_KEY = 'BO13tsTjl2y_vuX84DIzUbbWUgndqDKnvi7CF-9kkeK5ZBjeTRck4m5X8zKFLgN_-8erCil_UC4Ei1tE5fgmM-M';

function _urlB64ToUint8Array(b64) {
  const pad = b64 + '==='.slice((b64.length + 3) % 4);
  const raw = atob(pad.replace(/-/g, '+').replace(/_/g, '/'));
  return Uint8Array.from([...raw].map(c => c.charCodeAt(0)));
}

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(clients.claim()));

// SW 更新で subscription endpoint が変わった場合に自動再登録
self.addEventListener('pushsubscriptionchange', e => {
  e.waitUntil((async () => {
    const sub = await self.registration.pushManager.subscribe({
      userVisibleOnly: true,
      applicationServerKey: _urlB64ToUint8Array(VAPID_PUBLIC_KEY),
    });
    const json = sub.toJSON();
    await fetch(GAS_URL, {
      method: 'POST',
      body: JSON.stringify({ action: 'subscribe', endpoint: json.endpoint, p256dh: json.keys.p256dh, auth: json.keys.auth }),
      headers: { 'Content-Type': 'text/plain' },
    });
  })());
});

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

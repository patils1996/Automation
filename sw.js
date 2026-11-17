const CACHE_NAME = 'bpcl-dashboard-v21';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './pwa-install.js',
  './styles.css',
  './ui-refresh.css',
  './ui-refresh.js',
  './app.js',
  './trend.js',
  './comparison.js',
  './new_logic_dashboard.js',
  './outlets_status.js',
  './whatsapp_alerts.js',
  './whatsapp_config.js',
  './day_monitoring_config.js',
  './sample_data.js',
  './xlsx.full.min.js',
  './chart.umd.min.js',
  './bpcl_emblem_official.png',
  './bpcl_logo_card.png',
  './sidebar_skyline.png',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', (e) => {
  self.skipWaiting();
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS).catch(err => console.warn('Cache addAll failed:', err));
    })
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  e.respondWith(
    caches.match(e.request).then((res) => {
      return res || fetch(e.request);
    })
  );
});

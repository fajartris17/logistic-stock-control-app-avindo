const CACHE_NAME = 'avindo-logistic-basic-v1';
const ASSETS = ['/', '/static/css/style.css', '/static/js/app.js', '/static/assets/avindo_logo.svg'];
self.addEventListener('install', event => { event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))); });
self.addEventListener('fetch', event => { event.respondWith(fetch(event.request).catch(() => caches.match(event.request))); });

// Guarda as telas do app no celular para abrir rápido.
// Os dados (planilha) sempre vêm da internet, nunca do cache.
const CACHE = 'ajoduns-oya-v1';
const ARQUIVOS = ['./', './index.html', './config.js', './manifest.json', './icon-192.png', './icon-512.png', './apple-touch-icon.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARQUIVOS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(n => Promise.all(n.filter(x => x !== CACHE).map(x => caches.delete(x)))));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return; // planilha e fontes: direto da internet
  // Primeiro tenta a versão nova; sem internet, usa a guardada
  e.respondWith(
    fetch(e.request).then(r => {
      const copia = r.clone();
      caches.open(CACHE).then(c => c.put(e.request, copia));
      return r;
    }).catch(() => caches.match(e.request))
  );
});

// Cachea solo la cáscara de la app. Los datos nunca se guardan: siempre se piden a la API.
// Avisos push: el SDK de Firebase muestra solo las notificaciones que llegan con la app cerrada.
importScripts("config.js");
importScripts("https://www.gstatic.com/firebasejs/10.14.1/firebase-app-compat.js");
importScripts("https://www.gstatic.com/firebasejs/10.14.1/firebase-messaging-compat.js");
firebase.initializeApp(self.FIREBASE_CONFIG);
firebase.messaging();

const CACHE = "finanzas-v15";
const SHELL = ["./", "index.html", "config.js", "manifest.webmanifest", "icon-192.png", "icon-512.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
});

self.addEventListener("notificationclick", (e) => {
  e.notification.close();
  e.waitUntil(clients.matchAll({ type: "window", includeUncontrolled: true }).then((cs) => {
    if (cs.length) return cs[0].focus();
    return clients.openWindow("./");
  }));
});

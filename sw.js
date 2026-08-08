/* Service Worker - AtletTraning PWA
   Cache hanya untuk fail aplikasi (shell). SEMUA DATA sentiasa diambil
   terus (network only) daripada Google Sheet melalui Apps Script. */
const CACHE="atlettraning-v34";
const SHELL = ["./", "./index.html", "./manifest.json", "./logo.png", "./logo-192.png", "./logo-512.png", "./apple-touch-icon.png", "./favicon.png"];

self.addEventListener("install", (e) => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL)).catch(() => {}));
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    (async () => {
      const keys = await caches.keys();
      await Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)));
      await self.clients.claim();
    })(),
  );
});

self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  // Jangan sekali-kali cache panggilan data (Google Apps Script / Sheet).
  if (url.hostname.includes("google.com") || url.hostname.includes("googleusercontent.com")) return;
  if (e.request.method !== "GET") return;

  // Stale-while-revalidate: papar shell dari cache dengan serta-merta,
  // kemudian kemas kini cache di belakang tabir supaya app sentiasa terkini.
  e.respondWith(
    caches.match(e.request).then((cached) => {
      const rangkaian = fetch(e.request)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
          return res;
        })
        .catch(() => cached || caches.match("./index.html"));
      return cached || rangkaian;
    }),
  );
});

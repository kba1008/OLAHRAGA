/* Service Worker - AtletTraning PWA
   Cache hanya untuk fail aplikasi (shell). SEMUA DATA sentiasa diambil
   terus (network only) daripada Google Sheet melalui Apps Script. */
const CACHE = "atlettraning-v43";
const SHELL = ["./", "./index.html", "./manifest.json"];

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

  const htmlKah =
    e.request.mode === "navigate" ||
    (e.request.headers.get("accept") || "").includes("text/html");

  if (htmlKah) {
    // NETWORK FIRST untuk halaman app: sentiasa ambil versi terbaru dahulu,
    // cache hanya digunakan bila offline. Ini elak app lambat kemas kini.
    e.respondWith(
      fetch(e.request)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match(e.request).then((r) => r || caches.match("./index.html"))),
    );
    return;
  }

  // Aset lain: stale-while-revalidate.
  e.respondWith(
    caches.match(e.request).then((cached) => {
      const rangkaian = fetch(e.request)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(e.request, copy)).catch(() => {});
          return res;
        })
        .catch(() => cached);
      return cached || rangkaian;
    }),
  );
});

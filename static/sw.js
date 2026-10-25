// Service worker = a small script the browser runs in the background.
// Ours is deliberately SAFE: it never stores pages (so private admin data and
// fresh board content are never served stale). It only keeps style/script/icons
// and an "offline" page for when there's no internet.
const CACHE = "excuse-board-v1";          // bump this number if you change the list below
const OFFLINE_URL = "/offline";
const PRECACHE = [OFFLINE_URL, "/static/style.css", "/static/theme.js",
                  "/static/icons/icon-192.png"];

// On install: save the files above.
self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE).then((c) => c.addAll(PRECACHE)));
  self.skipWaiting();
});

// On activate: delete caches from older versions.
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const req = event.request;
  if (req.method !== "GET") return;          // never touch form submissions (POST)

  // Page visits: always go to the network. If offline, show the offline page.
  if (req.mode === "navigate") {
    event.respondWith(fetch(req).catch(() => caches.match(OFFLINE_URL)));
    return;
  }

  // Static files: try network first (so your CSS edits show up), fall back to cache.
  const url = new URL(req.url);
  if (url.origin === location.origin && url.pathname.startsWith("/static/")) {
    event.respondWith(
      fetch(req).then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy));
        return res;
      }).catch(() => caches.match(req))
    );
  }
});

// Service worker = a small script the browser runs in the background.
// Ours is deliberately SAFE: it never stores pages (so private admin data and
// fresh board content are never served stale). It only keeps style/script/icons
// and an "offline" page for when there's no internet.
const CACHE = "excuse-board-v2";          // bump this number if you change the list below
const OFFLINE_URL = "/offline";
const PRECACHE = [OFFLINE_URL, "/static/style.css", "/static/theme-init.js",
                  "/static/theme.js", "/static/app.js", "/static/icons/icon-192.png"];

// On install: save the files above. Each one is added separately, so a single
// failed file can't break the whole install.
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE).then((c) =>
      Promise.all(PRECACHE.map((url) => c.add(url).catch(() => {}))))
  );
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
    event.respondWith(
      fetch(req).catch(() => caches.match(OFFLINE_URL).then((r) => r || Response.error()))
    );
    return;
  }

  // Static files: try network first (so your CSS edits show up), fall back to cache.
  const url = new URL(req.url);
  if (url.origin === location.origin && url.pathname.startsWith("/static/")) {
    event.respondWith(
      fetch(req).then((res) => {
        // Only save good responses, so a 404/500 never gets cached.
        if (res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put(req, copy));
        }
        return res;
      }).catch(() => caches.match(req).then((r) => r || Response.error()))
    );
  }
});

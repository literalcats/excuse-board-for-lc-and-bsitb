// Small helpers that used to be inline code. (Inline code is blocked by our
// Content-Security-Policy, which is a good thing: it stops injected scripts.)

// 1) Register the service worker (only works on https or localhost)
if ("serviceWorker" in navigator) {
  navigator.serviceWorker.register("/sw.js").catch(() => {});
}

// 2) Forms with data-confirm="..." ask "are you sure?" before sending
document.addEventListener("submit", (e) => {
  const msg = e.target.dataset && e.target.dataset.confirm;
  if (msg && !window.confirm(msg)) e.preventDefault();
});

// 3) Buttons with data-print / data-reload
document.addEventListener("click", (e) => {
  if (e.target.closest("[data-print]")) window.print();
  if (e.target.closest("[data-reload]")) location.reload();
});

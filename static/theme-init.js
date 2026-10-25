// Runs in <head> before the page paints, so the right theme shows instantly.
// try/catch because some browsers block localStorage (e.g. private mode).
try {
  document.documentElement.dataset.theme = localStorage.getItem("theme") || "dark";
} catch (e) {
  document.documentElement.dataset.theme = "dark";
}

// Theme toggle: flips data-theme on <html> and remembers it in the browser.
// CSS reads [data-theme="..."] to decide which colors to use.
const btn = document.getElementById("theme-toggle");
function label() {
  if (btn) btn.textContent = document.documentElement.dataset.theme === "dark" ? "Light mode" : "Dark mode";
}
if (btn) {
  btn.addEventListener("click", () => {
    const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
    document.documentElement.dataset.theme = next;
    try { localStorage.setItem("theme", next); } catch (e) { /* storage blocked: ignore */ }
    label();
  });
}
label();

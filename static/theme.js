// Theme toggle: flips data-theme on <html> and remembers it in the browser.
// CSS reads [data-theme="..."] to decide which colors to use.
const btn = document.getElementById("theme-toggle");
function label() {
  btn.textContent = document.documentElement.dataset.theme === "dark" ? "Light mode" : "Dark mode";
}
btn.addEventListener("click", () => {
  const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
  document.documentElement.dataset.theme = next;
  localStorage.setItem("theme", next);
  label();
});
label();

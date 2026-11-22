// Light/dark mode toggle. Without a choice the site follows the system setting; a click
// stores the visitor's choice in this browser only. Loaded in <head> so the stored theme
// applies before the page paints (no flash).
(function () {
  var KEY = "formulaFormatter.theme";
  var root = document.documentElement;
  var media = window.matchMedia ? window.matchMedia("(prefers-color-scheme: dark)") : null;
  var SUN =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>';
  var MOON =
    '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>';

  var stored = null;
  try {
    stored = localStorage.getItem(KEY);
  } catch (e) {
    /* storage blocked: follow the system */
  }
  if (stored === "light" || stored === "dark") root.setAttribute("data-theme", stored);

  function current() {
    var t = root.getAttribute("data-theme");
    if (t) return t;
    return media && media.matches ? "dark" : "light";
  }

  function render(button) {
    var dark = current() === "dark";
    button.innerHTML = dark ? SUN : MOON;
    var label = dark ? "Switch to light mode" : "Switch to dark mode";
    button.setAttribute("aria-label", label);
    button.title = label;
  }

  document.addEventListener("DOMContentLoaded", function () {
    var button = document.createElement("button");
    button.type = "button";
    button.className = "theme-toggle";
    render(button);
    button.addEventListener("click", function () {
      var next = current() === "dark" ? "light" : "dark";
      root.setAttribute("data-theme", next);
      try {
        localStorage.setItem(KEY, next);
      } catch (e) {
        /* not remembered, still applied */
      }
      render(button);
    });
    if (media && media.addEventListener) media.addEventListener("change", function () { render(button); });
    document.body.appendChild(button);
  });
})();

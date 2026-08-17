(function loadImpiantiVisualAndFattoGuard() {
  "use strict";
  if (!document.querySelector('link[data-impianti-zebra-style]')) {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "./impianti-zebra-fatto-guard.css?v=20260817b";
    link.dataset.impiantiZebraStyle = "1";
    document.head.appendChild(link);
  }
  if (!document.querySelector('script[data-fatto-scroll-guard]')) {
    const script = document.createElement("script");
    script.src = "./fatto-scroll-guard.js?v=20260817a";
    script.defer = true;
    script.dataset.fattoScrollGuard = "1";
    document.head.appendChild(script);
  }
})();

(function loadDesktopFullscreenStyle() {
  "use strict";
  if (document.querySelector('link[data-desktop-fullscreen-style]')) return;
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = "./desktop-fullscreen.css?v=20260817a";
  link.dataset.desktopFullscreenStyle = "1";
  document.head.appendChild(link);
})();

(function loadEarlyErrorReporter() {
  "use strict";
  if (window.HeraClientErrorReporter?.installed || document.querySelector('script[data-client-error-reporter]')) return;
  const src = "./client-error-reporter.js?v=20260816a";
  if (document.readyState === "loading") {
    document.write(`<script src="${src}" data-client-error-reporter="1"><\/script>`);
    return;
  }
  const script = document.createElement("script");
  script.src = src;
  script.dataset.clientErrorReporter = "1";
  script.async = false;
  document.head.appendChild(script);
})();

(function loadPwaWhazzupContinuousCamera() {
  "use strict";
  const isNative = Boolean(
    window.Capacitor &&
    typeof window.Capacitor.isNativePlatform === "function" &&
    window.Capacitor.isNativePlatform()
  );
  if (isNative || document.querySelector('script[data-pwa-whazzup-camera]')) return;
  const script = document.createElement("script");
  script.src = "./pwa-whazzup-continuous-camera.js?v=20260817a";
  script.defer = true;
  script.dataset.pwaWhazzupCamera = "1";
  document.head.appendChild(script);
})();

(function loadSquadreCommessaThemes() {
  "use strict";
  if (!document.querySelector('link[data-squadre-commessa-themes]')) {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "./squadre-commessa-themes.css?v=20260817b";
    link.dataset.squadreCommessaThemes = "1";
    document.head.appendChild(link);
  }
  if (!document.querySelector('link[data-squadre-commessa-themes-visible-fix]')) {
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "./squadre-commessa-themes-visible-fix.css?v=20260817d";
    link.dataset.squadreCommessaThemesVisibleFix = "1";
    document.head.appendChild(link);
  }
  if (!document.querySelector('script[data-squadre-commessa-themes]')) {
    const script = document.createElement("script");
    script.src = "./squadre-commessa-themes.js?v=20260817b";
    script.defer = true;
    script.dataset.squadreCommessaThemes = "1";
    document.head.appendChild(script);
  }
})();

(function installLoadingHumor() {
  "use strict";

  const ROTATION_MS = 2500;
  const SLOW_NOTICE_MS = 8000;
  const messages = [
    ["🧑‍🌾", "Sto preparando il cantiere…"],
    ["🗺️", "Metto in ordine commesse e squadre…"],
    ["🤔", "Sto pensando… quasi fatto!"],
    ["🛠️", "Controllo che sia tutto al posto giusto…"],
    ["🌱", "Ancora un attimo, ci siamo…"]
  ];

  const controllers = new Map();

  function isVisible(surface) {
    return surface.closest(".hidden, [hidden], [aria-hidden='true']") === null;
  }

  function setMessage(surface, index) {
    const emoji = surface.querySelector("[data-loading-humor-emoji]");
    const message = surface.querySelector("[data-loading-humor-message]");
    const next = messages[index % messages.length];
    if (emoji) emoji.textContent = next[0];
    if (message) message.textContent = next[1];
  }

  function stop(surface) {
    const controller = controllers.get(surface);
    if (!controller) return;
    window.clearInterval(controller.rotationTimer);
    window.clearTimeout(controller.slowTimer);
    controllers.delete(surface);
  }

  function start(surface) {
    if (controllers.has(surface) || !isVisible(surface)) return;

    let index = 0;
    const slowNotice = surface.querySelector("[data-loading-humor-slow]");
    setMessage(surface, index);
    slowNotice?.classList.add("hidden");

    const rotationTimer = window.setInterval(() => {
      index = (index + 1) % messages.length;
      setMessage(surface, index);
    }, ROTATION_MS);

    const slowTimer = window.setTimeout(() => {
      if (!isVisible(surface)) return;
      window.clearInterval(rotationTimer);
      const emoji = surface.querySelector("[data-loading-humor-emoji]");
      const message = surface.querySelector("[data-loading-humor-message]");
      if (emoji) emoji.textContent = "🐌";
      if (message) message.textContent = "Connessione lenta… ma sto arrivando!";
      slowNotice?.classList.remove("hidden");
    }, SLOW_NOTICE_MS);

    controllers.set(surface, { rotationTimer, slowTimer });
  }

  function sync(surface) {
    if (isVisible(surface)) start(surface);
    else stop(surface);
  }

  document.querySelectorAll("[data-loading-humor-surface]").forEach((surface) => {
    surface.querySelector("[data-loading-humor-retry]")?.addEventListener("click", () => {
      window.location.reload();
    });

    const visibilityObserver = new MutationObserver(() => sync(surface));
    let observedNode = surface;
    while (observedNode && observedNode !== document.body) {
      visibilityObserver.observe(observedNode, {
        attributes: true,
        attributeFilter: ["class", "hidden", "aria-hidden"]
      });
      observedNode = observedNode.parentElement;
    }

    sync(surface);
  });
})();

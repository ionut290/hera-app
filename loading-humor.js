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
    return !surface.hidden
      && !surface.classList.contains("hidden")
      && surface.getAttribute("aria-hidden") !== "true";
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

    new MutationObserver(() => sync(surface)).observe(surface, {
      attributes: true,
      attributeFilter: ["class", "hidden", "aria-hidden"]
    });

    sync(surface);
  });
})();

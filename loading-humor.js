(() => {
  "use strict";

  const ROTATION_MS = 2500;
  const SLOW_NOTICE_MS = 5000;
  const LOGIN_FAILSAFE_MS = 4500;
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
      if (message) message.textContent = "Connessione lenta… apro comunque l’accesso.";
      slowNotice?.classList.remove("hidden");
    }, SLOW_NOTICE_MS);
    controllers.set(surface, { rotationTimer, slowTimer });
  }

  function sync(surface) {
    if (isVisible(surface)) start(surface);
    else stop(surface);
  }

  function revealLoginFailsafe() {
    let user = null;
    try { user = window.firebase?.auth?.()?.currentUser || null; } catch (_) {}
    if (user) return;

    const gate = document.getElementById("auth-gate");
    const loader = document.getElementById("app-startup-loading");
    if (!gate) return;

    if (loader) {
      loader.classList.add("hidden");
      loader.hidden = true;
      loader.setAttribute("aria-hidden", "true");
    }
    gate.hidden = false;
    gate.classList.remove("hidden");
    gate.style.removeProperty("display");
    gate.removeAttribute("aria-hidden");

    const feedback = document.getElementById("auth-email-feedback");
    if (feedback && !feedback.textContent.trim()) {
      feedback.textContent = "Accesso pronto. Inserisci email e password.";
    }
    window.__heraStartupLoginFailsafeUsed = true;
  }

  function install() {
    document.querySelectorAll("[data-loading-humor-surface]").forEach((surface) => {
      surface.querySelector("[data-loading-humor-retry]")?.addEventListener("click", () => window.location.reload());
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
    window.setTimeout(revealLoginFailsafe, LOGIN_FAILSAFE_MS);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();

(() => {
  "use strict";

  function addStyle(href, dataName) {
    if (document.querySelector(`link[data-${dataName}]`)) return;
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = href;
    link.setAttribute(`data-${dataName}`, "1");
    document.head.appendChild(link);
  }

  function addScript(src, dataName, defer = true) {
    if (document.querySelector(`script[data-${dataName}]`)) return;
    const script = document.createElement("script");
    script.src = src;
    script.defer = defer;
    script.setAttribute(`data-${dataName}`, "1");
    document.head.appendChild(script);
  }

  function loadOptionalRuntime() {
    addStyle("./impianti-zebra-fatto-guard.css?v=20260817b", "impianti-zebra-style");
    addStyle("./desktop-fullscreen.css?v=20260817a", "desktop-fullscreen-style");
    addStyle("./squadre-commessa-themes.css?v=20260817b", "squadre-commessa-themes");
    addStyle("./squadre-commessa-themes-visible-fix.css?v=20260817d", "squadre-commessa-themes-visible-fix");
    addScript("./fatto-scroll-guard.js?v=20260817a", "fatto-scroll-guard");
    addScript("./client-error-reporter.js?v=20260816a", "client-error-reporter");
    addScript("./app-error-monitor.js?v=20260824b", "app-error-monitor");
    addScript("./admin-error-center.js?v=20260824b", "admin-error-center");
    addStyle("./admin-error-center.css?v=20260824b", "admin-error-center-style");
    addScript("./squadre-commessa-themes.js?v=20260817b", "squadre-commessa-themes-script");
    addScript("./shared-pdf-attachments.js?v=20260821v2", "shared-whazzup-pdf");

    const isNative = Boolean(window.Capacitor?.isNativePlatform?.());
    if (!isNative) addScript("./pwa-whazzup-continuous-camera.js?v=20260817a", "pwa-whazzup-camera");
  }

  function scheduleOptionalRuntime() {
    const run = () => {
      if ("requestIdleCallback" in window) window.requestIdleCallback(loadOptionalRuntime, { timeout: 2500 });
      else window.setTimeout(loadOptionalRuntime, 1200);
    };
    if (document.readyState === "complete") run();
    else window.addEventListener("load", run, { once: true });
  }

  scheduleOptionalRuntime();
})();

(() => {
  "use strict";
  const REGION = "europe-west1";
  let attempts = 0;

  function attach() {
    const dialog = document.getElementById("hera-error-center-dialog");
    const actions = dialog?.querySelector?.(".hera-error-head-actions");
    if (!dialog || !actions) {
      if (attempts < 30) {
        attempts += 1;
        window.setTimeout(attach, 500);
      }
      return;
    }
    if (actions.querySelector("[data-error-reset]")) return;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn";
    button.dataset.errorReset = "1";
    button.textContent = "AZZERA";
    button.style.background = "#b91c1c";
    button.style.borderColor = "#b91c1c";
    button.style.color = "#fff";
    const closeButton = actions.querySelector("[data-error-close]");
    actions.insertBefore(button, closeButton || null);
    button.addEventListener("click", async () => {
      if (button.disabled) return;
      if (!window.confirm("Azzerare definitivamente tutti gli errori e i contatori del Centro errori? Questa operazione cancella solo la diagnostica del Centro errori e non modifica commesse, impianti, FATTO o WHAZZUP.")) return;
      const originalText = button.textContent;
      button.disabled = true;
      button.textContent = "AZZERAMENTO…";
      try {
        if (!window.firebase?.apps?.length || !window.firebase?.functions) throw new Error("Firebase non disponibile.");
        const callable = window.firebase.app().functions(REGION).httpsCallable("resetErrorCenter");
        const response = await callable({});
        const deleted = Math.max(0, Number(response?.data?.deletedGroups || 0));
        await window.HeraAdminErrorCenter?.refresh?.();
        window.alert(`✅ Centro errori azzerato. Eliminati ${deleted} gruppi di errore.`);
      } catch (error) {
        console.error("Azzeramento Centro errori fallito:", error);
        window.alert(`⚠️ Azzeramento non riuscito: ${error?.message || "errore sconosciuto"}`);
      } finally {
        button.disabled = false;
        button.textContent = originalText;
      }
    });
  }

  window.addEventListener("load", () => window.setTimeout(attach, 1800), { once: true });
})();

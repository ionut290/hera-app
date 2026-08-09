(function installPwaLoginForceUpdate() {
  "use strict";

  const CACHE_PREFIX = "varga-cantieri-";
  let updating = false;

  function setFeedback(message) {
    const feedback = document.getElementById("auth-email-feedback");
    if (feedback) feedback.textContent = message;
  }

  async function forcePwaUpdate() {
    if (updating) return;
    const button = document.getElementById("auth-force-update-btn");

    if (navigator.onLine === false) {
      setFeedback("Per aggiornare l’app è necessaria una connessione Internet.");
      return;
    }

    updating = true;
    if (button) {
      button.disabled = true;
      button.textContent = "AGGIORNAMENTO...";
    }
    setFeedback("Aggiornamento dell’app in corso. Attendi...");

    try {
      if ("serviceWorker" in navigator) {
        const registrations = await navigator.serviceWorker.getRegistrations();
        await Promise.allSettled(registrations.map(async (registration) => {
          await registration.update();
          registration.waiting?.postMessage({ type: "SKIP_WAITING" });
        }));
      }

      if ("caches" in window) {
        const keys = await caches.keys();
        await Promise.all(
          keys
            .filter((key) => key.startsWith(CACHE_PREFIX))
            .map((key) => caches.delete(key))
        );
      }

      const target = new URL(window.location.href);
      target.searchParams.set("pwa-refresh", String(Date.now()));
      setFeedback("Aggiornamento completato. Riapertura dell’app...");
      window.setTimeout(() => window.location.replace(target.href), 250);
    } catch (error) {
      console.error("Aggiornamento PWA non riuscito:", error);
      updating = false;
      if (button) {
        button.disabled = false;
        button.textContent = "🔄 AGGIORNA APP";
      }
      setFeedback("Aggiornamento non riuscito. Controlla Internet e riprova.");
    }
  }

  function initialize() {
    document.getElementById("auth-force-update-btn")
      ?.addEventListener("click", forcePwaUpdate);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();

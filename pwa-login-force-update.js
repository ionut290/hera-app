(function installPwaLoginForceUpdate() {
  "use strict";

  const CACHE_PREFIX = "varga-cantieri-";
  const APP_VERSION = "v116";
  let updating = false;

  function mountVersionBadges() {
    if (!document.getElementById("pwa-version-styles")) {
      const style = document.createElement("style");
      style.id = "pwa-version-styles";
      style.textContent = ".pwa-version-badge{display:flex;width:max-content;max-width:100%;align-items:center;justify-content:center;margin:.3rem auto .15rem;padding:.18rem .5rem;border:1px solid rgba(7,91,73,.18);border-radius:999px;background:rgba(237,249,245,.88);color:#45645d;font-size:.66rem;font-weight:800;line-height:1.15;white-space:nowrap}.pwa-version-badge.is-update{border-color:#d89b18;background:#fff7da;color:#805b00;cursor:pointer}#home-pwa-version{opacity:.82}";
      document.head.appendChild(style);
    }
    const add = (id, parent) => {
      if (!parent || document.getElementById(id)) return;
      const badge = document.createElement("button");
      badge.id = id;
      badge.type = "button";
      badge.className = "pwa-version-badge";
      badge.textContent = `PWA ${APP_VERSION} • Aggiornata`;
      badge.title = "Versione PWA installata";
      badge.addEventListener("click", () => badge.classList.contains("is-update") && forcePwaUpdate());
      parent.appendChild(badge);
    };
    add("auth-pwa-version", document.querySelector("#auth-gate .auth-gate-card"));
    add("home-pwa-version", document.getElementById("home-screen") || document.querySelector("main") || document.getElementById("app"));
  }

  function setUpdateAvailable() {
    ["auth-pwa-version", "home-pwa-version"].forEach((id) => {
      const badge = document.getElementById(id);
      if (!badge) return;
      badge.textContent = "⚠️ Aggiornamento PWA disponibile • Aggiorna";
      badge.classList.add("is-update");
      badge.title = "Premi per aggiornare la PWA";
    });
  }

  async function watchForUpdates() {
    if (!("serviceWorker" in navigator)) return;
    try {
      const registration = await navigator.serviceWorker.getRegistration();
      if (!registration) return;
      if (registration.waiting) setUpdateAvailable();
      registration.addEventListener("updatefound", () => {
        const worker = registration.installing;
        worker?.addEventListener("statechange", () => {
          if (worker.state === "installed" && navigator.serviceWorker.controller) setUpdateAvailable();
        });
      });
      if (navigator.onLine) await registration.update();
    } catch (error) {
      console.warn("Controllo aggiornamento PWA non riuscito:", error);
    }
  }

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
    mountVersionBadges();
    new MutationObserver(mountVersionBadges).observe(document.body, { childList: true, subtree: true });
    watchForUpdates();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }
})();

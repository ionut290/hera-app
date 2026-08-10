(function installAppUpdateButton() {
  "use strict";

  const PLAY_STORE_URL = "https://play.google.com/store/apps/details?id=it.vargacantieri.hera";
  let controllerReloaded = false;

  function isNativeAndroid() {
    return Boolean(
      window.Capacitor?.isNativePlatform?.()
      && window.Capacitor?.getPlatform?.() === "android"
    );
  }

  async function clearWebAppCaches() {
    if (!("caches" in window)) return;
    try {
      const cacheNames = await caches.keys();
      await Promise.all(
        cacheNames
          .filter((name) => name.startsWith("hera-app-shell-") || name.startsWith("varga-cantieri-shell-"))
          .map((name) => caches.delete(name))
      );
    } catch (error) {
      console.warn("Pulizia cache web non riuscita; proseguo con l'aggiornamento.", error);
    }
  }

  async function requestPwaUpdate({ reload = false } = {}) {
    if (isNativeAndroid() || !("serviceWorker" in navigator)) return false;
    try {
      const registration = await navigator.serviceWorker.getRegistration();
      await registration?.update?.();
      const waiting = registration?.waiting;
      if (waiting) waiting.postMessage({ type: "SKIP_WAITING" });
      if (reload) {
        await clearWebAppCaches();
        const refreshUrl = new URL(window.location.href);
        refreshUrl.searchParams.set("appRefresh", String(Date.now()));
        window.location.replace(refreshUrl.toString());
      }
      return true;
    } catch (error) {
      console.warn("Controllo aggiornamento web non riuscito.", error);
      return false;
    }
  }

  async function updateApplication(button) {
    button.disabled = true;
    button.setAttribute("aria-busy", "true");

    if (isNativeAndroid()) {
      window.location.assign(PLAY_STORE_URL);
      return;
    }

    await requestPwaUpdate({ reload: true });
    if (document.visibilityState === "visible") {
      const refreshUrl = new URL(window.location.href);
      refreshUrl.searchParams.set("appRefresh", String(Date.now()));
      window.location.replace(refreshUrl.toString());
    }
  }

  function installAutomaticPwaUpdate() {
    if (isNativeAndroid() || !("serviceWorker" in navigator)) return;

    navigator.serviceWorker.addEventListener("controllerchange", () => {
      if (controllerReloaded) return;
      controllerReloaded = true;
      window.location.reload();
    });

    const check = () => void requestPwaUpdate();
    window.addEventListener("online", check);
    window.addEventListener("pageshow", check);
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") check();
    });
    window.setTimeout(check, 1500);
  }

  function install() {
    const userButton = document.getElementById("user-toggle-btn");
    if (!userButton || document.getElementById("update-app-btn")) return;

    const style = document.createElement("style");
    style.textContent = `
      #update-app-btn {
        position: absolute;
        left: 40px;
        top: 50%;
        transform: translateY(-50%);
        width: 34px;
        height: 34px;
        padding: 0;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border: 1px solid #cfe0da;
        border-radius: 10px;
        background: linear-gradient(180deg, #fff 0%, #f3f8f6 100%);
        color: #184c3d;
        box-shadow: var(--shadow);
        font-size: 1rem;
        line-height: 1;
        cursor: pointer;
        z-index: 1;
      }
      #update-app-btn:focus-visible {
        outline: 3px solid rgba(37, 99, 235, 0.24);
        outline-offset: 2px;
      }
    `;
    document.head.appendChild(style);

    const button = document.createElement("button");
    button.id = "update-app-btn";
    button.type = "button";
    button.title = "Aggiorna app";
    button.setAttribute("aria-label", "Aggiorna app");
    button.textContent = "↻";
    button.addEventListener("click", () => void updateApplication(button));
    userButton.insertAdjacentElement("afterend", button);
  }

  installAutomaticPwaUpdate();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();

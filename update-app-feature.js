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
      await Promise.all(cacheNames.map((name) => caches.delete(name)));
    } catch (error) {
      console.warn("Pulizia cache web non riuscita; proseguo con l'aggiornamento.", error);
    }
  }

  function reloadWithoutCache() {
    const refreshUrl = new URL(window.location.href);
    refreshUrl.searchParams.set("appRefresh", String(Date.now()));
    window.location.replace(refreshUrl.toString());
  }

  async function requestPwaUpdate({ reload = false } = {}) {
    try {
      // La pulizia completa deve avvenire solo dopo un clic volontario su Refresh.
      // I controlli automatici aggiornano il Service Worker senza svuotare le cache.
      if (reload) await clearWebAppCaches();

      if ("serviceWorker" in navigator) {
        const registrations = await navigator.serviceWorker.getRegistrations();
        await Promise.all(registrations.map(async (registration) => {
          await registration.update().catch(() => null);
          registration.waiting?.postMessage({ type: "SKIP_WAITING" });
        }));
      }

      if (reload) reloadWithoutCache();
      return true;
    } catch (error) {
      console.warn("Controllo aggiornamento web non riuscito.", error);
      if (reload) reloadWithoutCache();
      return false;
    }
  }

  async function updateApplication(button) {
    if (!button || button.dataset.updateBusy === "1") return;
    button.dataset.updateBusy = "1";
    button.disabled = true;
    button.setAttribute("aria-busy", "true");

    try {
      const updated = await requestPwaUpdate({ reload: true });
      if (!updated && document.visibilityState === "visible") reloadWithoutCache();
    } finally {
      if (document.visibilityState === "visible") {
        button.disabled = false;
        button.removeAttribute("aria-busy");
        delete button.dataset.updateBusy;
      }
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

  function ensureStyles() {
    if (document.getElementById("pwa-update-compact-style")) return;
    const style = document.createElement("style");
    style.id = "pwa-update-compact-style";
    style.textContent = `
      #update-app-btn,
      #auth-update-pwa-btn {
        width: 30px;
        height: 30px;
        min-width: 30px;
        min-height: 30px;
        padding: 0;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border: 1px solid #cfe0da;
        border-radius: 9px;
        background: linear-gradient(180deg, #fff 0%, #f3f8f6 100%);
        color: #184c3d;
        box-shadow: 0 4px 10px rgba(15, 23, 42, 0.08);
        font-size: 0.95rem;
        font-weight: 800;
        line-height: 1;
        cursor: pointer;
        flex: 0 0 auto;
      }
      #update-app-btn span:last-child {
        position: absolute;
        width: 1px;
        height: 1px;
        padding: 0;
        margin: -1px;
        overflow: hidden;
        clip: rect(0, 0, 0, 0);
        white-space: nowrap;
        border: 0;
      }
      #auth-update-pwa-wrap {
        display: flex;
        justify-content: center;
        align-items: center;
        min-height: 30px;
        margin-top: 8px;
      }
      #auth-update-pwa-btn {
        width: auto;
        min-width: 30px;
        padding: 0 8px;
        gap: 5px;
        font-size: 0.75rem;
        font-weight: 700;
        white-space: nowrap;
      }
      #auth-update-pwa-btn .pwa-update-icon {
        font-size: 0.95rem;
        line-height: 1;
      }
      #update-app-btn:focus-visible,
      #auth-update-pwa-btn:focus-visible {
        outline: 3px solid rgba(37, 99, 235, 0.24);
        outline-offset: 2px;
      }
      @media (max-width: 420px) {
        #auth-update-pwa-btn {
          padding: 0 7px;
          font-size: 0.72rem;
        }
      }
    `;
    document.head.appendChild(style);
  }

  function bindUpdateButton(button) {
    if (!button || button.dataset.pwaUpdateBound === "1") return;
    button.dataset.pwaUpdateBound = "1";
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      void updateApplication(button);
    });
  }

  function installHomeButton() {
    const userButton = document.getElementById("user-toggle-btn");
    if (!userButton) return;

    let button = document.getElementById("update-app-btn");
    if (!button) {
      button = document.createElement("button");
      button.id = "update-app-btn";
      button.className = "update-app-btn";
      button.type = "button";
      button.innerHTML = '<span aria-hidden="true">↻</span><span>Aggiorna app</span>';
      userButton.insertAdjacentElement("afterend", button);
    }

    button.title = "Refresh: elimina cache e ricarica";
    button.setAttribute("aria-label", "Refresh: elimina cache e ricarica");
    bindUpdateButton(button);
  }

  function installLoginButton() {
    if (isNativeAndroid()) return;
    const card = document.querySelector("#auth-gate .auth-gate-card");
    if (!card || document.getElementById("auth-update-pwa-btn")) return;

    const wrap = document.createElement("div");
    wrap.id = "auth-update-pwa-wrap";

    const button = document.createElement("button");
    button.id = "auth-update-pwa-btn";
    button.type = "button";
    button.title = "Aggiorna la versione PWA";
    button.setAttribute("aria-label", "Aggiorna la versione PWA");
    button.innerHTML = '<span class="pwa-update-icon" aria-hidden="true">↻</span><span>Aggiorna PWA</span>';

    wrap.appendChild(button);
    card.appendChild(wrap);
    bindUpdateButton(button);
  }

  function install() {
    ensureStyles();
    installHomeButton();
    installLoginButton();
  }

  installAutomaticPwaUpdate();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();

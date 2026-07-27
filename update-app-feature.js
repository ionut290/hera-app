(function installAppUpdateButton() {
  "use strict";

  const PLAY_STORE_URL = "https://play.google.com/store/apps/details?id=it.vargacantieri.hera";

  function isNativeAndroid() {
    return Boolean(
      window.Capacitor?.isNativePlatform?.()
      && window.Capacitor?.getPlatform?.() === "android"
    );
  }

  async function updateApplication(button) {
    button.disabled = true;
    button.setAttribute("aria-busy", "true");

    if (isNativeAndroid()) {
      window.location.assign(PLAY_STORE_URL);
      return;
    }

    try {
      const registration = await navigator.serviceWorker?.getRegistration?.();
      await registration?.update?.();
    } catch (error) {
      console.warn("Controllo aggiornamento web non riuscito; ricarico la pagina corrente.", error);
    }
    window.location.reload();
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

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", install, { once: true });
  } else {
    install();
  }
})();

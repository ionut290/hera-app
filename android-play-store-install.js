(function installAndroidPlayStoreInstallButton() {
  "use strict";

  const PLAY_STORE_URL = "https://play.google.com/store/apps/details?id=it.vargacantieri.hera";

  function isNativeAndroidApp() {
    return Boolean(
      window.Capacitor?.isNativePlatform?.()
      && window.Capacitor?.getPlatform?.() === "android"
    );
  }

  function isAndroidDevice() {
    const userAgent = String(navigator.userAgent || "");
    const userAgentPlatform = String(navigator.userAgentData?.platform || "");
    return isNativeAndroidApp() || /Android/i.test(`${userAgent} ${userAgentPlatform}`);
  }

  function closeMenu() {
    if (typeof window.closeSideMenu === "function") window.closeSideMenu();
  }

  function openPlayStore() {
    const storeWindow = window.open(PLAY_STORE_URL, "_blank", "noopener,noreferrer");
    if (!storeWindow) window.location.href = PLAY_STORE_URL;
  }

  function handleAndroidInstallClick(event) {
    if (!isAndroidDevice()) return;

    event.preventDefault();
    event.stopImmediatePropagation();

    if (isNativeAndroidApp()) {
      alert("L'app Android risulta già installata sul dispositivo.");
      closeMenu();
      return;
    }

    const confirmed = window.confirm(
      "Vuoi aprire Google Play e installare l'app Android?\n\nOK = SÌ, INSTALLA\nAnnulla = NO"
    );
    closeMenu();
    if (confirmed) openPlayStore();
  }

  function bindInstallButton() {
    const button = document.getElementById("install-app-btn");
    if (!button || button.dataset.androidStoreInstallBound === "1") return;
    button.dataset.androidStoreInstallBound = "1";
    button.addEventListener("click", handleAndroidInstallClick, true);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bindInstallButton, { once: true });
  } else {
    bindInstallButton();
  }
})();

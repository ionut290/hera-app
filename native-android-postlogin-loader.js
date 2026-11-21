(function installAndroidPostloginNativeRuntimeLoader() {
  "use strict";
  if (window.__heraAndroidPostloginNativeRuntimeLoaderInstalled) return;
  window.__heraAndroidPostloginNativeRuntimeLoaderInstalled = true;

  const nativeAndroid = Boolean(
    window.Capacitor?.isNativePlatform?.() &&
    window.Capacitor?.getPlatform?.() === "android"
  );
  const scriptUrl = "native-android-runtime.js?v=20260803-whatsapp-early2";
  let loaded = false;
  let loading = false;

  function loadOriginalNativeRuntime() {
    if (loaded || loading) return;
    loading = true;
    const script = document.createElement("script");
    script.src = scriptUrl;
    script.async = false;
    script.dataset.heraAndroidNativeRuntime = "1";
    script.onload = () => { loaded = true; loading = false; };
    script.onerror = () => {
      loading = false;
      script.remove();
      console.warn("Runtime Android non disponibile: riprova dopo aver effettuato l'accesso.");
    };
    document.head.appendChild(script);
  }

  // Keep the existing early WhatsApp behavior intact on Web and iOS.
  if (!nativeAndroid) {
    loadOriginalNativeRuntime();
    return;
  }

  function startAfterAuth() {
    try {
      const auth = window.firebase?.auth?.();
      if (!auth || typeof auth.onAuthStateChanged !== "function") {
        console.warn("Autenticazione Android non ancora pronta: runtime GPS non avviato.");
        return;
      }
      auth.onAuthStateChanged((user) => {
        // No native GPS, geofence or notifications before an authenticated UID.
        if (!user?.uid) return;
        loadOriginalNativeRuntime();
      });
    } catch (error) {
      console.warn("Runtime nativo Android rimandato fino al login:", error?.code || "not-ready");
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startAfterAuth, { once: true });
  } else {
    startAfterAuth();
  }
})();

(function loadNativeRuntimeAfterAndroidLogin() {
  "use strict";
  if (window.__heraAndroidNativeRuntimeLoaderInstalled) return;
  window.__heraAndroidNativeRuntimeLoaderInstalled = true;

  const cap = window.Capacitor;
  const nativeAndroid = Boolean(
    cap?.isNativePlatform?.() && cap?.getPlatform?.() === "android"
  );
  const runtimeSrc = "native-android-runtime.js?v=20261121-native-auth-load1";
  let loadState = "idle";

  function loadUnmodifiedNativeRuntime() {
    if (loadState !== "idle") return;
    loadState = "loading";
    const script = document.createElement("script");
    script.src = runtimeSrc;
    script.async = false;
    script.onload = () => { loadState = "loaded"; };
    script.onerror = () => {
      loadState = "idle";
      console.warn("[varga-android] Runtime nativo non caricato: riprova.");
    };
    document.head.appendChild(script);
  }

  // Browser/PWA: preserve the original WhatsApp guard and normal web behavior.
  if (!nativeAndroid) {
    loadUnmodifiedNativeRuntime();
    return;
  }

  // The FATTO/Whazzup native runtime is immutable: do not execute its
  // automatic GPS/push startup on the login page, even if Capacitor's plugin
  // registry is read-only and the prelogin proxy cannot intercept calls.
  window.HeraNativeAndroid = {
    isNative: true,
    refreshLocation: async () => null,
    configureBackgroundLocation: async () => false,
    configureNotifications: async () => {},
    getLastPosition: () => null
  };

  let initialized = false;
  let authenticatedSessionSeen = false;
  let attempts = 0;
  function attachToAuth() {
    if (initialized) return;
    let auth = null;
    try { auth = window.firebase?.auth?.(); } catch (_) {}
    if (!auth || typeof auth.onAuthStateChanged !== "function") {
      if (++attempts < 80) window.setTimeout(attachToAuth, 250);
      else console.warn("[varga-android] Firebase Auth non pronto per il runtime nativo.");
      return;
    }
    initialized = true;
    auth.onAuthStateChanged((user) => {
      if (user?.uid) {
        authenticatedSessionSeen = true;
        loadUnmodifiedNativeRuntime();
      } else if (authenticatedSessionSeen) {
        // Destroy registered GPS/visibility handlers on sign-out. The next
        // login screen must be a fresh document without the native runtime.
        authenticatedSessionSeen = false;
        window.location.reload();
      }
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", attachToAuth, { once: true });
  } else {
    attachToAuth();
  }
})();
(function installAndroidPreloginPermissionGate() {
  "use strict";
  if (window.__heraAndroidPreloginPermissionGateInstalled) return;
  window.__heraAndroidPreloginPermissionGateInstalled = true;

  const cap = window.Capacitor;
  if (!cap?.isNativePlatform?.() || cap?.getPlatform?.() !== "android") return;

  let permissionStartupUid = "";
  function authenticatedUid() {
    try {
      return String(window.firebase?.auth?.().currentUser?.uid || "");
    } catch (_) {
      return "";
    }
  }
  function protectPlugin(name, blocked) {
    const original = cap.Plugins?.[name] || cap.registerPlugin?.(name);
    if (!original || !cap.Plugins) return;
    cap.Plugins[name] = new Proxy(original, {
      get(target, prop, receiver) {
        const method = Reflect.get(target, prop, receiver);
        if (!Object.prototype.hasOwnProperty.call(blocked, prop)) return method;
        if (typeof method !== "function") return method;
        return (...args) => {
          if (!authenticatedUid()) {
            const fallback = blocked[prop];
            return fallback instanceof Error
              ? Promise.reject(fallback)
              : Promise.resolve({ ...fallback });
          }
          return method.apply(target, args);
        };
      }
    });
  }

  const signedOutError = new Error("Accedi prima di usare la posizione Android.");
  protectPlugin("Geolocation", {
    checkPermissions: { location: "denied", coarseLocation: "denied" },
    requestPermissions: { location: "denied", coarseLocation: "denied" },
    getCurrentPosition: signedOutError,
    watchPosition: signedOutError
  });
  protectPlugin("HeraGeofence", {
    activate: { active: false, needsAuthentication: true }
  });
  protectPlugin("PushNotifications", {
    checkPermissions: { receive: "denied" },
    requestPermissions: { receive: "denied" },
    register: {}
  });

  function startPermissionsAfterLogin() {
    try {
      const auth = window.firebase?.auth?.();
      if (!auth?.onAuthStateChanged) return;
      // Existing authenticated sessions are initialized by the original
      // synchronous native runtime at DOMContentLoaded, so do not duplicate.
      permissionStartupUid = authenticatedUid();
      auth.onAuthStateChanged((user) => {
        if (!user?.uid) {
          permissionStartupUid = "";
          return;
        }
        if (permissionStartupUid === user.uid) return;
        permissionStartupUid = user.uid;
        const runtime = window.HeraNativeAndroid;
        if (!runtime) return;
        // Sequentially request permissions to avoid two Android GPS prompts
        // racing when the first login finishes.
        void Promise.resolve()
          .then(() => runtime.refreshLocation?.())
          .then(() => runtime.configureBackgroundLocation?.())
          .then(() => runtime.configureNotifications?.())
          .catch(error => console.warn("Permessi Android post-login:", error?.code || "unavailable"));
      });
    } catch (error) {
      console.warn("Attesa autenticazione prima dei permessi Android:", error?.code || "not-ready");
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", startPermissionsAfterLogin, { once: true });
  } else {
    startPermissionsAfterLogin();
  }
})();

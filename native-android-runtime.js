(function () {
  "use strict";

  const capacitorPlatform = (() => {
    try {
      return window.Capacitor && typeof window.Capacitor.getPlatform === "function"
        ? window.Capacitor.getPlatform()
        : "";
    } catch (_) {
      return "";
    }
  })();

  const nativeAndroid = Boolean(
    window.Capacitor &&
    (capacitorPlatform === "android" ||
      (typeof window.Capacitor.isNativePlatform === "function" &&
        window.Capacitor.isNativePlatform() &&
        /Android/i.test(navigator.userAgent)))
  );
  if (!nativeAndroid) return;

  window.__HERA_NATIVE_ANDROID__ = true;
  document.documentElement.dataset.heraPlatform = "android-native";
  document.documentElement.classList.add("hera-native-android");

  const plugin = (name) => window.Capacitor.Plugins?.[name]
    || window.Capacitor.registerPlugin?.(name)
    || null;

  const Geolocation = plugin("Geolocation");
  const PushNotifications = plugin("PushNotifications");
  const NOTIFICATION_KEY = "hera_notifications_configured_v1";
  let lastPermissionState = "unknown";

  function normalizeNativePosition(position) {
    if (!position || !position.coords) return position;
    return {
      coords: {
        latitude: Number(position.coords.latitude),
        longitude: Number(position.coords.longitude),
        accuracy: Number(position.coords.accuracy || 0),
        altitude: position.coords.altitude == null ? null : Number(position.coords.altitude),
        altitudeAccuracy: position.coords.altitudeAccuracy == null ? null : Number(position.coords.altitudeAccuracy),
        heading: position.coords.heading == null ? null : Number(position.coords.heading),
        speed: position.coords.speed == null ? null : Number(position.coords.speed)
      },
      timestamp: Number(position.timestamp || Date.now())
    };
  }

  function hasLocationPermission(permission) {
    return permission?.location === "granted" || permission?.coarseLocation === "granted";
  }

  async function refreshNativeLocation(options = {}) {
    if (!Geolocation) {
      window.dispatchEvent(new CustomEvent("hera:native-location-error", {
        detail: { code: "PLUGIN_MISSING", message: "Plugin posizione Android non disponibile." }
      }));
      return null;
    }

    try {
      let permission = await Geolocation.checkPermissions();
      if (!hasLocationPermission(permission) && options.requestPermission !== false) {
        permission = await Geolocation.requestPermissions({ permissions: ["location", "coarseLocation"] });
      }
      lastPermissionState = hasLocationPermission(permission) ? "granted" : "denied";
      if (!hasLocationPermission(permission)) {
        window.dispatchEvent(new CustomEvent("hera:native-location-error", {
          detail: { code: "PERMISSION_DENIED", message: "Permesso posizione Android non concesso." }
        }));
        return null;
      }

      const position = normalizeNativePosition(await Geolocation.getCurrentPosition({
        enableHighAccuracy: true,
        timeout: 20000,
        maximumAge: 30000
      }));
      window.__heraLastNativePosition = position;
      window.dispatchEvent(new CustomEvent("hera:native-location", { detail: position }));
      return position;
    } catch (error) {
      console.warn("Posizione nativa non disponibile:", error);
      window.dispatchEvent(new CustomEvent("hera:native-location-error", {
        detail: { code: error?.code || "LOCATION_ERROR", message: error?.message || "Posizione Android non disponibile." }
      }));
      return null;
    }
  }

  function installNativeGeolocationBridge() {
    const successAsync = (callback, value) => {
      if (typeof callback === "function") setTimeout(() => callback(value), 0);
    };
    const errorAsync = (callback, message, code = 1) => {
      if (typeof callback !== "function") return;
      setTimeout(() => callback({ code, message }), 0);
    };

    const activeWatches = new Map();
    let nextWatchId = 1;
    const bridge = {
      getCurrentPosition(success, error, options) {
        refreshNativeLocation({ requestPermission: true, ...(options || {}) }).then((position) => {
          if (position) successAsync(success, position);
          else errorAsync(error, "Permesso posizione Android non concesso o posizione non disponibile.");
        }).catch(() => errorAsync(error, "Posizione Android non disponibile."));
      },
      watchPosition(success, error) {
        const watchId = nextWatchId++;
        const update = () => refreshNativeLocation({ requestPermission: true }).then((position) => {
          if (position) successAsync(success, position);
          else errorAsync(error, "Posizione Android non disponibile.");
        });
        update();
        activeWatches.set(watchId, setInterval(update, 30000));
        return watchId;
      },
      clearWatch(watchId) {
        const timer = activeWatches.get(watchId);
        if (timer) clearInterval(timer);
        activeWatches.delete(watchId);
      }
    };

    window.__heraNativeGeolocationBridge = bridge;
    try {
      Object.defineProperty(navigator, "geolocation", {
        configurable: true,
        enumerable: true,
        value: bridge
      });
    } catch (error) {
      console.warn("Impossibile sostituire navigator.geolocation.", error);
    }
  }

  function suppressBrowserOnlyLocationWarning() {
    const fixWarning = () => {
      const platform = document.getElementById("map-location-warning-platform");
      if (platform && /Chrome|browser/i.test(platform.textContent || "")) {
        platform.textContent = "App Android • Posizione nativa";
      }
      const warning = document.getElementById("map-location-warning");
      if (warning && window.__heraLastNativePosition) {
        warning.classList.add("hidden");
      }
    };
    fixWarning();
    new MutationObserver(fixWarning).observe(document.documentElement, {
      subtree: true,
      childList: true,
      characterData: true,
      attributes: true,
      attributeFilter: ["class"]
    });

    document.addEventListener("click", (event) => {
      const button = event.target?.closest?.("#map-enable-location-btn, #map-retry-location-btn");
      if (!button) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      refreshNativeLocation({ requestPermission: true });
    }, true);
  }

  async function disableWebServiceWorkerInNativeApp() {
    try {
      if ("serviceWorker" in navigator) {
        const registrations = await navigator.serviceWorker.getRegistrations();
        await Promise.all(registrations.map((registration) => registration.unregister()));
      }
      if ("caches" in window) {
        const keys = await caches.keys();
        await Promise.all(keys.map((key) => caches.delete(key)));
      }
    } catch (error) {
      console.warn("Pulizia cache web Android non riuscita:", error);
    }
  }

  async function configureNotificationsOnce() {
    if (!PushNotifications || localStorage.getItem(NOTIFICATION_KEY) === "done") return;
    try {
      let permission = await PushNotifications.checkPermissions();
      if (permission.receive === "prompt" || permission.receive === "prompt-with-rationale") {
        permission = await PushNotifications.requestPermissions();
      }
      if (permission.receive !== "granted") return;
      await PushNotifications.register();
      localStorage.setItem(NOTIFICATION_KEY, "done");
    } catch (error) {
      console.warn("Notifiche native non configurate:", error);
    }
  }

  window.HeraNativeAndroid = {
    refreshLocation: refreshNativeLocation,
    configureNotifications: configureNotificationsOnce,
    getLastPosition: () => window.__heraLastNativePosition || null,
    getPermissionState: () => lastPermissionState,
    isNative: true
  };

  installNativeGeolocationBridge();
  disableWebServiceWorkerInNativeApp();

  const start = () => {
    suppressBrowserOnlyLocationWarning();
    refreshNativeLocation({ requestPermission: true });
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", start, { once: true });
  } else {
    start();
  }

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") refreshNativeLocation({ requestPermission: false });
  });
})();
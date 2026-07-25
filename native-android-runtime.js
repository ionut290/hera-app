(function () {
  "use strict";

  const nativeAndroid = Boolean(
    window.Capacitor &&
    typeof window.Capacitor.isNativePlatform === "function" &&
    window.Capacitor.isNativePlatform() &&
    window.Capacitor.getPlatform() === "android"
  );
  if (!nativeAndroid) return;

  const plugin = (name) => window.Capacitor.Plugins?.[name]
    || window.Capacitor.registerPlugin?.(name)
    || null;

  const Geolocation = plugin("Geolocation");
  const PushNotifications = plugin("PushNotifications");
  const NOTIFICATION_KEY = "hera_notifications_configured_v1";

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

  async function refreshNativeLocation() {
    if (!Geolocation) return null;
    try {
      let permission = await Geolocation.checkPermissions();
      if (permission.location !== "granted" && permission.coarseLocation !== "granted") {
        permission = await Geolocation.requestPermissions({ permissions: ["location", "coarseLocation"] });
      }
      if (permission.location !== "granted" && permission.coarseLocation !== "granted") return null;

      const position = normalizeNativePosition(await Geolocation.getCurrentPosition({
        enableHighAccuracy: true,
        timeout: 15000,
        maximumAge: 60000
      }));
      window.__heraLastNativePosition = position;
      window.dispatchEvent(new CustomEvent("hera:native-location", { detail: position }));
      return position;
    } catch (error) {
      console.warn("Posizione nativa non disponibile:", error);
      return null;
    }
  }

  function installNativeGeolocationBridge() {
    const original = navigator.geolocation;
    const successAsync = (callback, value) => {
      if (typeof callback === "function") setTimeout(() => callback(value), 0);
    };
    const errorAsync = (callback, message) => {
      if (typeof callback !== "function") return;
      setTimeout(() => callback({ code: 1, message }), 0);
    };

    const bridge = {
      getCurrentPosition(success, error) {
        refreshNativeLocation().then((position) => {
          if (position) successAsync(success, position);
          else errorAsync(error, "Permesso posizione Android non concesso o posizione non disponibile.");
        }).catch(() => errorAsync(error, "Posizione Android non disponibile."));
      },
      watchPosition(success, error) {
        const watchId = Date.now() + Math.floor(Math.random() * 1000);
        refreshNativeLocation().then((position) => {
          if (position) successAsync(success, position);
          else errorAsync(error, "Permesso posizione Android non concesso o posizione non disponibile.");
        });
        return watchId;
      },
      clearWatch() {}
    };

    try {
      Object.defineProperty(navigator, "geolocation", {
        configurable: true,
        enumerable: true,
        value: bridge
      });
    } catch (error) {
      console.warn("Impossibile sostituire navigator.geolocation; uso bridge compatibile.", error);
      window.__heraNativeGeolocationBridge = bridge;
      if (!original) navigator.geolocation = bridge;
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
    isNative: true
  };

  installNativeGeolocationBridge();

  const start = () => {
    refreshNativeLocation();
    configureNotificationsOnce();
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", start, { once: true });
  } else {
    start();
  }

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") refreshNativeLocation();
  });
})();

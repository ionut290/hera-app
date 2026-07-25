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

  async function refreshNativeLocation() {
    if (!Geolocation) return null;
    try {
      let permission = await Geolocation.checkPermissions();
      if (permission.location !== "granted" && permission.coarseLocation !== "granted") {
        permission = await Geolocation.requestPermissions({ permissions: ["location", "coarseLocation"] });
      }
      if (permission.location !== "granted" && permission.coarseLocation !== "granted") return null;

      const position = await Geolocation.getCurrentPosition({
        enableHighAccuracy: true,
        timeout: 15000,
        maximumAge: 60000
      });
      window.__heraLastNativePosition = position;
      window.dispatchEvent(new CustomEvent("hera:native-location", { detail: position }));
      return position;
    } catch (error) {
      console.warn("Posizione nativa non disponibile:", error);
      return null;
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
    getLastPosition: () => window.__heraLastNativePosition || null
  };

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

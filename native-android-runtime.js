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
  const HeraGeofence = plugin("HeraGeofence");
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

  async function saveNativePushToken(token) {
    if (!token) return;
    localStorage.setItem("heraPushFcmToken", token);
    window.dispatchEvent(new CustomEvent("hera:native-push-token", { detail: { token } }));

    try {
      const user = window.firebase?.auth?.().currentUser;
      if (!user) return;
      await window.firebase.firestore().collection("platformUsers").doc(user.uid).set({
        pushToken: token,
        pushPlatform: "android-native",
        pushTokenUpdatedAt: window.firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
    } catch (error) {
      console.warn("Token push nativo non salvato nel profilo:", error);
    }
  }

  async function configureNotificationsOnce() {
    if (!PushNotifications) return;
    try {
      await PushNotifications.removeAllListeners();
      await PushNotifications.addListener("registration", ({ value }) => saveNativePushToken(value));
      await PushNotifications.addListener("registrationError", (error) => {
        console.warn("Registrazione notifiche Android fallita:", error);
      });
      await PushNotifications.addListener("pushNotificationActionPerformed", ({ notification }) => {
        window.dispatchEvent(new CustomEvent("hera:native-notification-opened", { detail: notification }));
      });

      await PushNotifications.createChannel({
        id: "hera_operational_updates",
        name: "Aggiornamenti operativi",
        description: "Notifiche quando un impianto viene segnato come FATTO",
        importance: 5,
        visibility: 1,
        vibration: true
      });

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

  function persistPushTokenAfterLogin() {
    try {
      const auth = window.firebase?.auth?.();
      if (!auth || typeof auth.onAuthStateChanged !== "function") return;
      auth.onAuthStateChanged((user) => {
        const token = localStorage.getItem("heraPushFcmToken");
        if (user && token) saveNativePushToken(token);
      });
    } catch (error) {
      console.warn("Associazione notifiche Android al login non disponibile:", error);
    }
  }

  async function configureBackgroundLocation() {
    if (!HeraGeofence) return false;
    try {
      const status = await HeraGeofence.activate();
      if (status?.needsBackgroundSettings) {
        console.warn(status.message || "Autorizzazione posizione in background richiesta dalle impostazioni Android.");
        window.dispatchEvent(new CustomEvent("hera:background-location-settings-required", {
          detail: status
        }));
        return false;
      }
      return Boolean(status?.active);
    } catch (error) {
      console.warn("Posizione Android in background non attivata:", error);
      return false;
    }
  }

  window.HeraNativeAndroid = {
    refreshLocation: refreshNativeLocation,
    configureBackgroundLocation,
    configureNotifications: configureNotificationsOnce,
    getLastPosition: () => window.__heraLastNativePosition || null,
    isNative: true
  };

  installNativeGeolocationBridge();

  const start = () => {
    persistPushTokenAfterLogin();
    refreshNativeLocation();
    configureBackgroundLocation();
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

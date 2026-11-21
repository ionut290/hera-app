(function installEarlyPwaWhatsAppGuard() {
  "use strict";

  if (window.__heraEarlyPwaWhatsAppGuardInstalled) return;
  window.__heraEarlyPwaWhatsAppGuardInstalled = true;

  const isNativeAndroid = Boolean(
    window.Capacitor &&
    typeof window.Capacitor.isNativePlatform === "function" &&
    window.Capacitor.isNativePlatform() &&
    window.Capacitor.getPlatform() === "android"
  );
  if (isNativeAndroid) return;

  function isWhatsAppWebLink(value) {
    if (typeof value !== "string") return false;
    try {
      const url = new URL(value, window.location.href);
      const host = url.hostname.toLowerCase();
      return host === "api.whatsapp.com" || host === "wa.me" || host === "www.wa.me" || host === "web.whatsapp.com";
    } catch (_) {
      return false;
    }
  }

  function toWhatsAppScheme(value) {
    if (typeof value !== "string") return "";
    if (value.startsWith("whatsapp://")) return value;
    try {
      const url = new URL(value, window.location.href);
      const params = new URLSearchParams();
      const text = url.searchParams.get("text") || url.searchParams.get("message") || "";
      const phoneFromQuery = url.searchParams.get("phone") || url.searchParams.get("send") || "";
      const phoneFromPath = /(?:^|\/)\+?([0-9]{6,})(?:\/|$)/.exec(url.pathname)?.[1] || "";
      const phone = String(phoneFromQuery || phoneFromPath).replace(/\D/g, "");
      if (phone) params.set("phone", phone);
      if (text) params.set("text", text);
      if (!phone && !text) return "";
      return `whatsapp://send?${params.toString()}`;
    } catch (_) {
      return "";
    }
  }

  function openInstalledWhatsApp(value) {
    const directUrl = toWhatsAppScheme(value);
    if (!directUrl) {
      window.alert("Messaggio WhatsApp non disponibile.");
      return false;
    }
    window.location.href = directUrl;
    window.setTimeout(() => {
      if (document.visibilityState === "visible") {
        window.alert("WhatsApp non è installato o non può essere aperto su questo dispositivo.");
      }
    }, 1800);
    return true;
  }

  const originalOpen = window.open.bind(window);
  window.open = function heraEarlyWhatsAppOpen(url, target, features) {
    if (isWhatsAppWebLink(url)) {
      openInstalledWhatsApp(url);
      return null;
    }
    return originalOpen(url, target, features);
  };

  document.addEventListener("click", (event) => {
    const anchor = event.target?.closest?.("a[href]");
    if (!anchor || !isWhatsAppWebLink(anchor.href)) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    openInstalledWhatsApp(anchor.href);
  }, true);
})();

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
  const nativeLocationWatchers = new Map();
  let nativeLocationWatchCounter = 0;

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

  function publishNativePosition(position) {
    const normalized = normalizeNativePosition(position);
    if (!normalized?.coords) return null;
    window.__heraLastNativePosition = normalized;
    window.dispatchEvent(new CustomEvent("hera:native-location", { detail: normalized }));
    return normalized;
  }

  async function refreshNativeLocation() {
    if (!Geolocation) return null;
    try {
      let permission = await Geolocation.checkPermissions();
      if (permission.location !== "granted" && permission.coarseLocation !== "granted") {
        permission = await Geolocation.requestPermissions({ permissions: ["location", "coarseLocation"] });
      }
      if (permission.location !== "granted" && permission.coarseLocation !== "granted") return null;

      const position = publishNativePosition(await Geolocation.getCurrentPosition({
        enableHighAccuracy: true,
        timeout: 15000,
        maximumAge: 0
      }));
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
      watchPosition(success, error, options = {}) {
        const watchId = ++nativeLocationWatchCounter;
        const watcher = { nativeId: null, cancelled: false, fallbackTimer: null };
        nativeLocationWatchers.set(watchId, watcher);

        if (!Geolocation || typeof Geolocation.watchPosition !== "function") {
          const refresh = () => refreshNativeLocation().then((position) => {
            if (position) successAsync(success, position);
            else errorAsync(error, "Posizione Android non disponibile.");
          });
          refresh();
          watcher.fallbackTimer = setInterval(refresh, 15000);
          return watchId;
        }

        const startNativeWatch = async () => {
          const initialPosition = await refreshNativeLocation();
          const activeBeforeStart = nativeLocationWatchers.get(watchId);
          if (!activeBeforeStart || activeBeforeStart.cancelled) return null;
          if (!initialPosition) throw new Error("Permesso posizione Android non concesso o posizione non disponibile.");
          successAsync(success, initialPosition);
          return Geolocation.watchPosition({
            enableHighAccuracy: options.enableHighAccuracy !== false,
            timeout: Number(options.timeout || 15000),
            maximumAge: Math.min(Number(options.maximumAge || 0), 10000)
          }, (position, watchError) => {
            const activeWatcher = nativeLocationWatchers.get(watchId);
            if (!activeWatcher || activeWatcher.cancelled) return;
            if (watchError || !position) {
              errorAsync(error, watchError?.message || "Posizione Android non disponibile.");
              return;
            }
            const normalized = publishNativePosition(position);
            if (normalized) successAsync(success, normalized);
          });
        };

        Promise.resolve(startNativeWatch()).then((nativeId) => {
          const activeWatcher = nativeLocationWatchers.get(watchId);
          if (!activeWatcher || activeWatcher.cancelled) {
            if (nativeId != null) Geolocation.clearWatch({ id: nativeId }).catch(() => {});
            return;
          }
          activeWatcher.nativeId = nativeId;
        }).catch((watchError) => {
          nativeLocationWatchers.delete(watchId);
          errorAsync(error, watchError?.message || "Posizione Android non disponibile.");
        });
        return watchId;
      },
      clearWatch(watchId) {
        const watcher = nativeLocationWatchers.get(watchId);
        if (!watcher) return;
        watcher.cancelled = true;
        if (watcher.fallbackTimer) clearInterval(watcher.fallbackTimer);
        if (watcher.nativeId != null && Geolocation?.clearWatch) {
          Geolocation.clearWatch({ id: watcher.nativeId }).catch(() => {});
        }
        nativeLocationWatchers.delete(watchId);
      }
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

  function loadAndroidWhazzupPhotoOrderFix() {
    if (document.querySelector('script[data-hera-android-whazzup-photo-order="1"]')) return;
    const script = document.createElement("script");
    script.src = "android-whazzup-photo-order.js?v=20260814a";
    script.dataset.heraAndroidWhazzupPhotoOrder = "1";
    script.async = false;
    document.head.appendChild(script);
  }

  installNativeGeolocationBridge();

  const start = () => {
    loadAndroidWhazzupPhotoOrderFix();
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

(function installNotificationCompatibilityShim() {
  "use strict";

  const AUTO_ENABLE_NOTIFICATIONS_KEY = "heraAutoEnableNotifications";
  const PUSH_TOKEN_KEY = "heraPushFcmToken";

  // Le notifiche push web sono intenzionalmente disattivate.
  // Il Centro Notifiche interno e la relativa cronologia restano disponibili.
  function installWebPushDisabledGuard() {
    try {
      localStorage.setItem(AUTO_ENABLE_NOTIFICATIONS_KEY, "false");
      localStorage.removeItem(PUSH_TOKEN_KEY);
    } catch (_) {}

    const storagePrototype = window.Storage?.prototype;
    if (storagePrototype && !storagePrototype.setItem?.__heraWebPushDisabledGuard) {
      const originalSetItem = storagePrototype.setItem;
      const guardedSetItem = function guardedSetItem(key, value) {
        if (String(key) === AUTO_ENABLE_NOTIFICATIONS_KEY) {
          return originalSetItem.call(this, key, "false");
        }
        return originalSetItem.call(this, key, value);
      };
      try {
        Object.defineProperty(guardedSetItem, "__heraWebPushDisabledGuard", {
          value: true,
          configurable: false,
          enumerable: false,
          writable: false
        });
        storagePrototype.setItem = guardedSetItem;
      } catch (_) {}
    }

    // Neutralizza solo il token push web. Auth, Firestore e flussi operativi restano invariati.
    try {
      if (window.firebase?.messaging && typeof window.firebase.messaging === "function") {
        const messaging = window.firebase.messaging();
        if (messaging && !messaging.__heraWebPushDisabled) {
          messaging.getToken = async () => "";
          Object.defineProperty(messaging, "__heraWebPushDisabled", {
            value: true,
            configurable: false,
            enumerable: false,
            writable: false
          });
        }
      }
    } catch (_) {}

    window.HERA_WEB_PUSH_DISABLED = true;
  }

  installWebPushDisabledGuard();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", installWebPushDisabledGuard, { once: true });
  }
  window.setTimeout(installWebPushDisabledGuard, 250);
  window.setTimeout(installWebPushDisabledGuard, 1000);

  const LAST_NOTIFICATION_KEY = "hera_last_received_notification_v1";
  const NOTIFICATION_HISTORY_KEY = "hera_notification_history_v1";
  const MAX_HISTORY_ITEMS = 100;

  function normalizeNotification(raw) {
    const data = raw?.data || raw?.notification?.data || {};
    return {
      id: String(raw?.id || data.notificationId || data.id || Date.now()),
      title: String(raw?.title || raw?.notification?.title || data.title || "Varga Cantieri"),
      body: String(data.fullMessage || raw?.fullMessage || raw?.body || raw?.notification?.body || data.body || "Hai ricevuto una nuova notifica."),
      destination: String(data.destination || data.page || data.route || data.url || "home"),
      receivedAt: Number(raw?.receivedAt || Date.now()),
      data
    };
  }

  function loadHistory() {
    try {
      const rows = JSON.parse(localStorage.getItem(NOTIFICATION_HISTORY_KEY) || "[]");
      return Array.isArray(rows) ? rows : [];
    } catch (_) {
      return [];
    }
  }

  function archiveNotifications(rows) {
    const merged = new Map(loadHistory().map((item) => [String(item.id), item]));
    (Array.isArray(rows) ? rows : [rows]).filter(Boolean).forEach((raw) => {
      const item = normalizeNotification(raw);
      merged.set(item.id, { ...merged.get(item.id), ...item });
    });
    const history = [...merged.values()]
      .sort((a, b) => Number(b.receivedAt || 0) - Number(a.receivedAt || 0))
      .slice(0, MAX_HISTORY_ITEMS);
    try {
      localStorage.setItem(NOTIFICATION_HISTORY_KEY, JSON.stringify(history));
    } catch (_) {}
    return history;
  }

  function saveNotification(raw) {
    const notification = normalizeNotification(raw);
    try {
      localStorage.setItem(LAST_NOTIFICATION_KEY, JSON.stringify(notification));
    } catch (_) {}
    archiveNotifications(notification);
    return notification;
  }

  function removeLegacyUi() {
    document.getElementById("received-notification-dialog")?.remove();
    document.querySelectorAll("#notification-inbox-btn:not([data-central-notification-bell='1'])").forEach((node) => node.remove());
  }

  function openCentralNotificationCenter(raw) {
    if (raw) saveNotification(raw);
    removeLegacyUi();
    if (window.HeraNotificationCenter?.open) {
      window.HeraNotificationCenter.open();
      return;
    }
    setTimeout(() => window.HeraNotificationCenter?.open?.(), 250);
  }

  function showNotificationInbox() {
    openCentralNotificationCenter();
  }

  function openDestination(raw) {
    const notification = normalizeNotification(raw);
    window.dispatchEvent(new CustomEvent("hera:open-notification-destination", {
      detail: { destination: notification.destination, notification }
    }));
  }

  function showNotificationReader(raw) {
    const notification = saveNotification(raw);
    openCentralNotificationCenter(notification);
  }

  function installNativeListeners() {
    const nativeAndroid = Boolean(
      window.Capacitor
      && typeof window.Capacitor.isNativePlatform === "function"
      && window.Capacitor.isNativePlatform()
      && window.Capacitor.getPlatform?.() === "android"
    );
    if (!nativeAndroid) return;
    const push = window.Capacitor?.Plugins?.PushNotifications || window.Capacitor?.registerPlugin?.("PushNotifications");
    if (!push?.addListener) return;

    push.addListener("pushNotificationReceived", (notification) => {
      saveNotification(notification);
      window.HeraNotificationCenter?.open?.();
    }).catch?.(() => {});

    push.addListener("pushNotificationActionPerformed", (action) => {
      const notification = saveNotification(action?.notification || action);
      openDestination(notification);
      window.HeraNotificationCenter?.open?.();
    }).catch?.(() => {});
  }

  function installWebListeners() {
    if (!("serviceWorker" in navigator)) return;
    navigator.serviceWorker.addEventListener("message", (event) => {
      if (event.data?.type !== "HERA_NOTIFICATION_READ") return;
      const notification = saveNotification(event.data.notification || event.data);
      openDestination(notification);
      window.HeraNotificationCenter?.open?.();
    });

    const params = new URLSearchParams(location.search);
    if (params.get("notification") !== "open") return;
    params.delete("notification");
    const cleanUrl = `${location.pathname}${params.toString() ? `?${params}` : ""}${location.hash}`;
    history.replaceState({}, "", cleanUrl);
    setTimeout(() => window.HeraNotificationCenter?.open?.(), 500);
  }

  function start() {
    installWebPushDisabledGuard();
    removeLegacyUi();
    installNativeListeners();
    installWebListeners();
  }

  window.HeraNotificationReader = {
    show: showNotificationReader,
    archive: archiveNotifications,
    showAll: showNotificationInbox,
    openLast() {
      try {
        const saved = JSON.parse(localStorage.getItem(LAST_NOTIFICATION_KEY) || "null");
        if (saved) showNotificationReader(saved);
        else showNotificationInbox();
      } catch (_) {
        showNotificationInbox();
      }
    }
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
})();
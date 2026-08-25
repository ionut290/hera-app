(function installNotificationCompatibilityShim() {
  "use strict";

  const CORRECTED_DEFAULT_PUSH_PUBLIC_VAPID_KEY = "BLWYWSC_rEbfAoOnOaO6JYhaYVBCa7IDZaN-2cGMt6uqUYLWwl6mKq8hng9V5B5GPVUOlgjLPLhqz2KvdsuJUoA";

  function isValidPushPublicVapidKey(value) {
    const key = String(value || "").trim();
    if (!key || !/^[A-Za-z0-9_-]+$/.test(key)) return false;
    try {
      const normalized = key.replace(/-/g, "+").replace(/_/g, "/");
      const padded = normalized + "=".repeat((4 - (normalized.length % 4)) % 4);
      const bytes = Uint8Array.from(atob(padded), (char) => char.charCodeAt(0));
      return bytes.length === 65 && bytes[0] === 4;
    } catch (_) {
      return false;
    }
  }

  let storedPushPublicVapidKey = "";
  try {
    storedPushPublicVapidKey = String(localStorage.getItem("heraPushPublicVapidKey") || "").trim();
    if (storedPushPublicVapidKey && !isValidPushPublicVapidKey(storedPushPublicVapidKey)) {
      localStorage.removeItem("heraPushPublicVapidKey");
      storedPushPublicVapidKey = "";
    }
  } catch (_) {}

  const configuredPushPublicVapidKey = [
    window.HERA_PUSH_PUBLIC_VAPID_KEY,
    document.querySelector('meta[name="hera-push-vapid-key"]')?.content,
    storedPushPublicVapidKey
  ].find(isValidPushPublicVapidKey);

  window.HERA_PUSH_PUBLIC_VAPID_KEY = configuredPushPublicVapidKey || CORRECTED_DEFAULT_PUSH_PUBLIC_VAPID_KEY;

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
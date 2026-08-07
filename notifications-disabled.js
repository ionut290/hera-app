(() => {
  "use strict";

  if (window.__HERA_NOTIFICATIONS_DISABLED__) return;
  window.__HERA_NOTIFICATIONS_DISABLED__ = true;

  const BLOCKED_COLLECTIONS = new Set([
    "notifications",
    "userAlerts",
    "appNotifications",
    "userAlertAcknowledgements"
  ]);

  function canonicalPath(value) {
    if (!value) return "";
    if (typeof value === "string") return value;
    if (typeof value.canonicalString === "function") {
      try { return String(value.canonicalString() || ""); } catch (_) {}
    }
    if (typeof value.toArray === "function") {
      try {
        const parts = value.toArray();
        if (Array.isArray(parts)) return parts.join("/");
      } catch (_) {}
    }
    if (Array.isArray(value.segments)) return value.segments.join("/");
    if (Array.isArray(value._segments)) return value._segments.join("/");
    return "";
  }

  function firestorePath(target) {
    const candidates = [
      target?.path,
      target?._query?.path,
      target?._query?._path,
      target?._key?.path,
      target?._delegate?._query?.path,
      target?._delegate?._query?._path
    ];
    for (const candidate of candidates) {
      const path = canonicalPath(candidate).replace(/^\/+|\/+$/g, "");
      if (path) return path;
    }
    return "";
  }

  function isBlocked(target) {
    const path = firestorePath(target);
    const collection = path.split("/")[0] || "";
    return BLOCKED_COLLECTIONS.has(collection);
  }

  function emptyQuerySnapshot(query) {
    const docs = Object.freeze([]);
    const metadata = Object.freeze({ fromCache: true, hasPendingWrites: false });
    const snapshot = {
      query,
      docs,
      size: 0,
      empty: true,
      metadata,
      forEach() {},
      docChanges() { return []; },
      isEqual(other) { return other === snapshot; }
    };
    return Object.freeze(snapshot);
  }

  function emptyDocumentSnapshot(ref) {
    const metadata = Object.freeze({ fromCache: true, hasPendingWrites: false });
    return Object.freeze({
      id: String(ref?.id || ""),
      ref,
      exists: false,
      metadata,
      data: () => undefined,
      get: () => undefined,
      isEqual(other) { return other?.ref === ref && other?.exists === false; }
    });
  }

  function parseObserver(argsLike) {
    const args = Array.from(argsLike || []);
    let index = 0;
    if (args[0] && typeof args[0] === "object" && Object.prototype.hasOwnProperty.call(args[0], "includeMetadataChanges")) index = 1;
    const candidate = args[index];
    if (typeof candidate === "function") return { next: candidate, context: null };
    if (candidate && typeof candidate === "object" && typeof candidate.next === "function") {
      return { next: candidate.next, context: candidate };
    }
    return { next: null, context: null };
  }

  function deliverEmpty(target, argsLike, documentMode = false) {
    const observer = parseObserver(argsLike);
    if (observer.next) {
      const value = documentMode ? emptyDocumentSnapshot(target) : emptyQuerySnapshot(target);
      queueMicrotask(() => {
        try { observer.next.call(observer.context || undefined, value); } catch (error) { console.warn("Notification shutdown callback ignored", error); }
      });
    }
    return () => {};
  }

  function installFirestoreGuards() {
    const firestore = window.firebase?.firestore;
    const QueryProto = firestore?.Query?.prototype;
    const DocumentProto = firestore?.DocumentReference?.prototype;
    const CollectionProto = firestore?.CollectionReference?.prototype;
    if (!QueryProto || !DocumentProto || !CollectionProto) return false;

    if (typeof QueryProto.onSnapshot === "function" && !QueryProto.onSnapshot.__heraNotificationsDisabled) {
      const original = QueryProto.onSnapshot;
      const wrapped = function() {
        if (isBlocked(this)) return deliverEmpty(this, arguments, false);
        return original.apply(this, arguments);
      };
      Object.defineProperty(wrapped, "__heraNotificationsDisabled", { value: true });
      Object.defineProperty(wrapped, "__heraNotificationsOriginal", { value: original });
      QueryProto.onSnapshot = wrapped;
    }

    if (typeof QueryProto.get === "function" && !QueryProto.get.__heraNotificationsDisabled) {
      const original = QueryProto.get;
      const wrapped = function() {
        if (isBlocked(this)) return Promise.resolve(emptyQuerySnapshot(this));
        return original.apply(this, arguments);
      };
      Object.defineProperty(wrapped, "__heraNotificationsDisabled", { value: true });
      QueryProto.get = wrapped;
    }

    if (typeof DocumentProto.onSnapshot === "function" && !DocumentProto.onSnapshot.__heraNotificationsDisabled) {
      const original = DocumentProto.onSnapshot;
      const wrapped = function() {
        if (isBlocked(this)) return deliverEmpty(this, arguments, true);
        return original.apply(this, arguments);
      };
      Object.defineProperty(wrapped, "__heraNotificationsDisabled", { value: true });
      DocumentProto.onSnapshot = wrapped;
    }

    if (typeof DocumentProto.get === "function" && !DocumentProto.get.__heraNotificationsDisabled) {
      const original = DocumentProto.get;
      const wrapped = function() {
        if (isBlocked(this)) return Promise.resolve(emptyDocumentSnapshot(this));
        return original.apply(this, arguments);
      };
      Object.defineProperty(wrapped, "__heraNotificationsDisabled", { value: true });
      DocumentProto.get = wrapped;
    }

    for (const method of ["set", "update", "delete"]) {
      if (typeof DocumentProto[method] !== "function" || DocumentProto[method].__heraNotificationsDisabled) continue;
      const original = DocumentProto[method];
      const wrapped = function() {
        if (isBlocked(this)) return Promise.resolve(null);
        return original.apply(this, arguments);
      };
      Object.defineProperty(wrapped, "__heraNotificationsDisabled", { value: true });
      DocumentProto[method] = wrapped;
    }

    if (typeof CollectionProto.add === "function" && !CollectionProto.add.__heraNotificationsDisabled) {
      const original = CollectionProto.add;
      const wrapped = function() {
        if (isBlocked(this)) return Promise.resolve(null);
        return original.apply(this, arguments);
      };
      Object.defineProperty(wrapped, "__heraNotificationsDisabled", { value: true });
      CollectionProto.add = wrapped;
    }

    return true;
  }

  function removeNotificationUi() {
    [
      "notification-inbox-btn",
      "notification-center",
      "open-panel-notifiche",
      "panel-notifiche",
      "user-alert-modal",
      "notification-doc-viewer-modal",
      "pwa-notification-status",
      "enable-notifications-btn",
      "test-notification-btn",
      "today-alerts-btn"
    ].forEach((id) => document.getElementById(id)?.remove());
    document.querySelectorAll(".notification-toast,.notification-bell,.notification-bell-badge").forEach((node) => node.remove());
  }

  function clearLegacyNotificationState() {
    try {
      [
        "hera_notification_outbox_v1",
        "heraNotificationPermissionRequested",
        "heraPushToken",
        "heraPushTokenUpdatedAt"
      ].forEach((key) => localStorage.removeItem(key));
    } catch (_) {}
  }

  async function unsubscribeBrowserPush() {
    try {
      if (!navigator.serviceWorker || !window.PushManager) return;
      const registration = await navigator.serviceWorker.ready;
      const subscription = await registration?.pushManager?.getSubscription?.();
      if (subscription) await subscription.unsubscribe();
    } catch (error) {
      console.warn("Disiscrizione push non riuscita:", error);
    }
  }

  if (!installFirestoreGuards()) {
    let attempts = 0;
    const timer = setInterval(() => {
      attempts += 1;
      if (installFirestoreGuards() || attempts >= 80) clearInterval(timer);
    }, 50);
  }

  clearLegacyNotificationState();
  void unsubscribeBrowserPush();
  removeNotificationUi();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", removeNotificationUi, { once: true });
  }
  setTimeout(removeNotificationUi, 500);
  setTimeout(removeNotificationUi, 2000);

  window.HeraNotificationsDisabled = Object.freeze({
    installed: true,
    blockedCollections: Object.freeze(Array.from(BLOCKED_COLLECTIONS))
  });

  console.info("[HERA] Sistema notifiche disattivato completamente.");
})();

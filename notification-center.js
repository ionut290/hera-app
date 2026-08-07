(() => {
  "use strict";
  // Sistema notifiche rimosso. Questo file resta come stub temporaneo perché
  // index.html lo referenzia ancora; non apre listener, non crea UI e non scrive dati.
  window.HeraNotificationCenter = undefined;

  try {
    // Evita che il codice legacy di app.js ricrei automaticamente una nuova
    // sottoscrizione FCM quando il browser aveva già concesso il permesso.
    localStorage.setItem("heraPushFcmToken", "DISABLED");
    localStorage.removeItem("hera_notification_outbox_v1");
  } catch (_) {}

  const cleanup = () => {
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
  };

  const unsubscribePush = async () => {
    try {
      if (!navigator.serviceWorker || !window.PushManager) return;
      const registration = await navigator.serviceWorker.ready;
      const subscription = await registration?.pushManager?.getSubscription?.();
      if (subscription) await subscription.unsubscribe();
    } catch (error) {
      console.warn("Disiscrizione push non riuscita:", error);
    }
  };

  cleanup();
  void unsubscribePush();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", cleanup, { once: true });
  }
})();

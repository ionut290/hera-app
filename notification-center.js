(() => {
  "use strict";
  // Sistema notifiche rimosso. Questo file resta come stub temporaneo perché
  // index.html lo referenzia ancora; non apre listener, non crea UI e non scrive dati.
  window.HeraNotificationCenter = undefined;

  const cleanup = () => {
    [
      "notification-inbox-btn",
      "notification-center",
      "open-panel-notifiche",
      "panel-notifiche",
      "user-alert-modal",
      "notification-doc-viewer-modal"
    ].forEach((id) => document.getElementById(id)?.remove());
  };

  cleanup();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", cleanup, { once: true });
  }
})();

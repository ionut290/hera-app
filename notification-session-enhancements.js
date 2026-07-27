(function installNotificationAndSessionEnhancements() {
  "use strict";

  const LAST_NOTIFICATION_KEY = "hera_last_received_notification_v1";

  function isNativeAndroid() {
    return Boolean(
      window.Capacitor &&
      typeof window.Capacitor.isNativePlatform === "function" &&
      window.Capacitor.isNativePlatform() &&
      window.Capacitor.getPlatform?.() === "android"
    );
  }

  async function enablePersistentFirebaseSession() {
    if (!window.firebase || typeof firebase.auth !== "function") return;
    try {
      await firebase.auth().setPersistence(firebase.auth.Auth.Persistence.LOCAL);
      localStorage.setItem("hera_auth_persistence", "local");
    } catch (error) {
      console.warn("Persistenza login non configurata:", error);
    }
  }

  function normalizeNotification(raw) {
    const data = raw?.data || raw?.notification?.data || {};
    return {
      id: String(raw?.id || data.notificationId || data.id || Date.now()),
      title: String(raw?.title || raw?.notification?.title || data.title || "Varga Cantieri"),
      body: String(raw?.body || raw?.notification?.body || data.body || "Hai ricevuto una nuova notifica."),
      destination: String(data.destination || data.page || data.route || data.url || "home"),
      receivedAt: Date.now(),
      data
    };
  }

  function saveNotification(notification) {
    try {
      localStorage.setItem(LAST_NOTIFICATION_KEY, JSON.stringify(notification));
    } catch (_) {}
  }

  function ensureDialog() {
    let dialog = document.getElementById("received-notification-dialog");
    if (dialog) return dialog;

    dialog = document.createElement("dialog");
    dialog.id = "received-notification-dialog";
    dialog.setAttribute("aria-labelledby", "received-notification-title");
    dialog.innerHTML = `
      <form method="dialog" style="min-width:min(92vw,520px);max-width:520px;padding:0;border:0">
        <header style="padding:18px 20px 10px">
          <p style="margin:0 0 5px;font-size:12px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;opacity:.65">Notifica ricevuta</p>
          <h2 id="received-notification-title" style="margin:0;font-size:22px">Varga Cantieri</h2>
        </header>
        <div style="padding:8px 20px 18px">
          <p id="received-notification-body" style="white-space:pre-wrap;line-height:1.5;margin:0"></p>
        </div>
        <footer style="display:flex;gap:10px;justify-content:flex-end;padding:14px 20px 20px;flex-wrap:wrap">
          <button id="received-notification-close" class="btn" value="cancel" type="submit">LEGGI E CHIUDI</button>
          <button id="received-notification-open" class="btn btn-primary" value="default" type="button">APRI NELL’APP</button>
        </footer>
      </form>`;
    document.body.appendChild(dialog);
    return dialog;
  }

  function openDestination(notification) {
    const destination = notification?.destination || "home";
    window.dispatchEvent(new CustomEvent("hera:open-notification-destination", {
      detail: { destination, notification }
    }));

    const selectors = {
      notifiche: "#open-panel-notifiche",
      programmazione: "#open-panel-programmazione",
      ore: "#open-hours-btn",
      segnalazioni: "#open-segnalazioni-btn",
      pos: "#open-pos-btn",
      home: "#back-to-home-btn"
    };
    const selector = selectors[destination] || (destination.startsWith("#") ? destination : null);
    const target = selector ? document.querySelector(selector) : null;
    if (target && typeof target.click === "function") target.click();
    else window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function showNotificationReader(raw) {
    const notification = normalizeNotification(raw);
    saveNotification(notification);
    const dialog = ensureDialog();
    dialog.querySelector("#received-notification-title").textContent = notification.title;
    dialog.querySelector("#received-notification-body").textContent = notification.body;
    const openButton = dialog.querySelector("#received-notification-open");
    openButton.onclick = () => {
      dialog.close();
      openDestination(notification);
    };
    if (typeof dialog.showModal === "function") {
      if (!dialog.open) dialog.showModal();
    } else {
      alert(`${notification.title}\n\n${notification.body}`);
    }
  }

  function installAndroidNotificationListeners() {
    if (!isNativeAndroid()) return;
    const push = window.Capacitor?.Plugins?.PushNotifications || window.Capacitor?.registerPlugin?.("PushNotifications");
    if (!push?.addListener) return;

    push.addListener("pushNotificationReceived", (notification) => {
      showNotificationReader(notification);
    }).catch?.(() => {});

    push.addListener("pushNotificationActionPerformed", (action) => {
      showNotificationReader(action?.notification || action);
    }).catch?.(() => {});
  }

  function installWebNotificationListeners() {
    if (!("serviceWorker" in navigator)) return;
    navigator.serviceWorker.addEventListener("message", (event) => {
      if (event.data?.type === "HERA_NOTIFICATION_READ") {
        showNotificationReader(event.data.notification || event.data);
      }
    });

    const params = new URLSearchParams(location.search);
    if (params.get("notification") === "open") {
      try {
        const saved = JSON.parse(localStorage.getItem(LAST_NOTIFICATION_KEY) || "null");
        if (saved) setTimeout(() => showNotificationReader(saved), 500);
      } catch (_) {}
      params.delete("notification");
      const cleanUrl = `${location.pathname}${params.toString() ? `?${params}` : ""}${location.hash}`;
      history.replaceState({}, "", cleanUrl);
    }
  }

  function start() {
    enablePersistentFirebaseSession();
    installAndroidNotificationListeners();
    installWebNotificationListeners();
  }

  window.HeraNotificationReader = {
    show: showNotificationReader,
    openLast() {
      try {
        const saved = JSON.parse(localStorage.getItem(LAST_NOTIFICATION_KEY) || "null");
        if (saved) showNotificationReader(saved);
      } catch (_) {}
    }
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
})();

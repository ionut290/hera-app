/* Varga Cantieri: promemoria iOS settimanali anche con applicazione chiusa.
 * Nessun tracciamento in background e nessuna timbratura automatica.
 * Il suono rispetta Silenzioso / Full Immersion e i permessi iOS.
 */
(() => {
  "use strict";
  if (window.VargaTimbraturaIOS) return;

  const BASE_ID = 620100, ACTION = "VARGA_TIMBRATURA";
  const DAYS = [0, 1, 2, 3, 4, 5, 6];
  const APP_STORE = "https://apps.apple.com/it/app/easy-box/id1612276209";
  const SHORTCUT = "shortcuts://run-shortcut?name=Apri%20Easy-Box";
  let listenerReady = false, pending = null, busy = false;

  const isNativeIOS = () => {
    const c = window.Capacitor;
    return !!(c?.isNativePlatform?.() && c?.getPlatform?.() === "ios");
  };
  function getPlugin() {
    const c = window.Capacitor;
    if (!isNativeIOS() || !c?.isPluginAvailable?.("LocalNotifications")) return null;
    return c.Plugins?.LocalNotifications || c.registerPlugin?.("LocalNotifications") || null;
  }
  const allIDs = () => DAYS.flatMap(day => [BASE_ID + day * 2, BASE_ID + day * 2 + 1]);
  function makeNotifications(config) {
    if (!config || !config.iosScheduled) return [];
    const selected = Array.isArray(config.days) ? config.days : [];
    const notices = [];
    for (const day of selected) {
      if (!Number.isInteger(day) || day < 0 || day > 6) continue;
      for (const [which, suffix, key] of [["arrival", 0, "arrivalFrom"], ["return", 1, "returnFrom"]]) {
        const m = /^([01]\d|2[0-3]):([0-5]\d)$/.exec(config[key] || "");
        if (!m) continue;
        notices.push({
          id: BASE_ID + day * 2 + suffix,
          title: which === "arrival" ? "🔔 Ricorda di timbrare l'ENTRATA" : "🔔 Ricorda di timbrare l'USCITA",
          body: "Apri Varga Cantieri e vai alla timbratura con Easy-Box. Questo avviso non registra la presenza.",
          sound: "default",
          actionTypeId: ACTION,
          threadIdentifier: "varga-timbratura",
          // JS weekday 0=Sun, iOS calendar weekday 1=Sun.
          schedule: { on: { weekday: day + 1, hour: Number(m[1]), minute: Number(m[2]), second: 0 }, repeats: true },
          extra: { feature: "varga-timbratura", which, destination: "easy-box" }
        });
      }
    }
    return notices;
  }

  async function init() {
    const local = getPlugin();
    if (!local || listenerReady) return false;
    listenerReady = true;
    try {
      await local.registerActionTypes({
        types: [{ id: ACTION, actions: [{ id: "OPEN_EASYBOX", title: "Apri timbratura", foreground: true }] }]
      });
      await local.addListener("localNotificationActionPerformed", payload => {
        // Capacitor may surface native Core Location notifications in userInfo.
        // Never act on an unrelated app notification.
        const info = payload?.notification?.extra ||
          payload?.notification?.userInfo || payload?.notification || {};
        if (info.feature !== "varga-timbratura") return;
        // Do not open another app if the user DISMISSES the alert.
        const action = payload?.actionId;
        if (action !== "tap" && action !== "OPEN_EASYBOX") return;
        pending = info.which === "return" ? "return" : "arrival";
        // A real notification tap is required for the optional Easy-Box redirect.
        window.dispatchEvent(new CustomEvent("varga:apri-timbratura", {
          detail: { which: pending, source: "notification-tap", tapped: true }
        }));
      });
    } catch (error) {
      listenerReady = false;
      console.warn("Notifiche timbratura iOS: ascolto non disponibile", error);
    }
    return listenerReady;
  }

  async function sync(config, options = {}) {
    if (!isNativeIOS()) return { supported: false, reason: "Solo nell'app iPhone installata." };
    const plugin = getPlugin();
    if (!plugin) return { supported: false, reason: "Aggiorna Varga Cantieri: notifiche iPhone non incluse in questa versione." };
    if (busy) return { supported: true, busy: true };
    busy = true;
    try {
      await init();
      // Cancel ONLY our own stable IDs: never cancel reminders of other features.
      await plugin.cancel({ notifications: allIDs().map(id => ({ id })) });
      if (!config?.iosScheduled) return { supported: true, scheduled: 0 };
      let authorization = await plugin.checkPermissions();
      if (authorization?.display !== "granted" && options.prompt) {
        authorization = await plugin.requestPermissions();
      }
      if (authorization?.display !== "granted") {
        return { supported: true, scheduled: 0, permissionDenied: true };
      }
      const notifications = makeNotifications(config);
      if (!notifications.length) return { supported: true, scheduled: 0 };
      await plugin.schedule({ notifications });
      return { supported: true, scheduled: notifications.length };
    } catch (error) {
      console.warn("Pianificazione avvisi iPhone non riuscita", error);
      return { supported: true, scheduled: 0, error: String(error?.message || error) };
    } finally {
      busy = false;
    }
  }

  function openEasyBoxStore() {
    // Store page is verified. We do not claim this is an undocumented app deep link.
    window.open(APP_STORE, "_blank", "noopener");
  }
  async function launchEasyBoxFromNotification() {
    if (!isNativeIOS()) return { opened: false, reason: "Solo sull'iPhone con app installata." };
    const geofence = window.Capacitor?.Plugins?.VargaGeofence ||
      window.Capacitor?.registerPlugin?.("VargaGeofence");
    if (!geofence?.openEasyBoxShortcut) {
      return { opened: false, reason: "Serve una nuova versione iOS con apertura nativa Easy-Box." };
    }
    try {
      const result = await geofence.openEasyBoxShortcut();
      return { opened: result?.opened === true, reason: result?.opened ? "" :
        "Non riesco ad aprire Comandi Rapidi. Verifica che sia installato." };
    } catch (error) {
      return { opened: false, reason: String(error?.message || error) };
    }
  }

  function openEasyBoxShortcut() {
    // Requires a user-created Apple Shortcut named "Apri Easy-Box".
    // No automatic app launch without a user action or a documented app deep link.
    window.location.href = SHORTCUT;
  }

  window.VargaTimbraturaIOS = {
    init, sync, makeNotifications, allIDs, isNativeIOS, launchEasyBoxFromNotification,
    openEasyBoxStore, openEasyBoxShortcut,
    get pending() { return pending; },
    clearPending() { pending = null; },
    APP_STORE
  };
})();
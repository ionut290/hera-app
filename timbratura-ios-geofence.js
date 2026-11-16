/* Opt-in geofence iOS: permission Always, GPS-triggered reminder, no background web timers. */
(() => {
  "use strict";
  if (window.VargaTimbraturaGeofence) return;
  function plugin() {
    const c = window.Capacitor;
    if (!c?.isNativePlatform?.() || c.getPlatform?.() !== "ios") return null;
    return c.Plugins?.VargaGeofence || c.registerPlugin?.("VargaGeofence") || null;
  }
  const isAvailable = () => !!plugin();

  function showPending(pending) {
    const which = pending?.which;
    if (which !== "arrival" && which !== "return") return;
    window.dispatchEvent(new CustomEvent("varga:apri-timbratura", { detail: { which, source: "geofence" } }));
  }

  async function status({ consumePending = false } = {}) {
    const p = plugin();
    if (!p?.status) return { supported: false, reason: "Richiede la nuova build nativa di Varga Cantieri per iPhone." };
    try {
      const result = await p.status({ consumePending });
      if (consumePending) showPending(result?.pending);
      return result;
    } catch (error) {
      return { supported: false, error: String(error?.message || error) };
    }
  }

  async function sync(config) {
    const p = plugin();
    if (!p?.configure || !p?.disable) {
      return { supported: false, reason: "Aggiorna l'app Varga Cantieri da TestFlight per il GPS con app chiusa." };
    }
    try {
      if (!config?.iosGeofence) return await p.disable();
      const numbers = ["lat", "lon", "radius"];
      if (numbers.some(k => !Number.isFinite(Number(config[k])) || config[k] === "")) {
        return { supported: true, error: "Coordinate o raggio mancanti." };
      }
      if (Number(config.radius) < 100 || Number(config.radius) > 2000) {
        return { supported: true, error: "Per il geofencing iPhone usa un raggio tra 100 e 2000 metri, consigliati 150–300." };
      }
      return await p.configure({
        latitude: Number(config.lat),
        longitude: Number(config.lon),
        radius: Number(config.radius),
        days: config.days,
        arrivalFrom: config.arrivalFrom,
        arrivalTo: config.arrivalTo,
        returnFrom: config.returnFrom,
        returnTo: config.returnTo
      });
    } catch (error) {
      return { supported: true, error: String(error?.message || error) };
    }
  }

  // When a native notification is tapped while an existing WebView resumes.
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) void status({ consumePending: true });
  });
  window.VargaTimbraturaGeofence = { isAvailable, sync, status };
})();
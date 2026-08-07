(() => {
  "use strict";

  if (window.HeraOfflineFirstRuntime?.installed) return;

  const STATE = {
    online: navigator.onLine !== false,
    syncing: false,
    lastSyncAt: null,
    lastError: null
  };

  const STYLE_ID = "hera-offline-first-style";
  const BANNER_ID = "hera-offline-first-banner";
  const SYNC_TIMEOUT_MS = 15000;

  function ensureStyle() {
    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
      #${BANNER_ID}{position:fixed;left:50%;bottom:14px;transform:translateX(-50%);z-index:100000;max-width:min(92vw,620px);padding:10px 14px;border-radius:12px;background:#1f2328;color:#fff;box-shadow:0 8px 30px rgba(0,0,0,.24);font:600 14px/1.3 system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;text-align:center;display:none}
      #${BANNER_ID}[data-state="offline"],#${BANNER_ID}[data-state="syncing"],#${BANNER_ID}[data-state="error"]{display:block}
      #${BANNER_ID}[data-state="online"]{display:block;opacity:.92}
    `;
    document.head.appendChild(style);
  }

  function ensureBanner() {
    ensureStyle();
    let banner = document.getElementById(BANNER_ID);
    if (!banner) {
      banner = document.createElement("div");
      banner.id = BANNER_ID;
      banner.setAttribute("role", "status");
      banner.setAttribute("aria-live", "polite");
      document.body.appendChild(banner);
    }
    return banner;
  }

  let hideTimer = null;
  function render(state, message, autoHideMs = 0) {
    const banner = ensureBanner();
    banner.dataset.state = state;
    banner.textContent = message;
    clearTimeout(hideTimer);
    if (autoHideMs > 0) {
      hideTimer = setTimeout(() => {
        if (navigator.onLine !== false) {
          banner.style.display = "none";
          banner.dataset.state = "";
        }
      }, autoHideMs);
    } else {
      banner.style.display = "";
    }
  }

  function getDb() {
    try {
      if (!window.firebase || typeof firebase.firestore !== "function") return null;
      return firebase.firestore();
    } catch (_) {
      return null;
    }
  }

  function withTimeout(promise, timeoutMs) {
    return Promise.race([
      promise,
      new Promise((_, reject) => setTimeout(() => reject(new Error("sync-timeout")), timeoutMs))
    ]);
  }

  async function syncPendingWrites() {
    if (navigator.onLine === false || STATE.syncing) return false;
    const db = getDb();
    if (!db) return false;

    STATE.syncing = true;
    STATE.lastError = null;
    render("syncing", "🔄 Connessione ripristinata. Sincronizzo le modifiche salvate offline…");

    try {
      if (typeof db.enableNetwork === "function") {
        await db.enableNetwork().catch(() => {});
      }
      if (typeof db.waitForPendingWrites === "function") {
        await withTimeout(db.waitForPendingWrites(), SYNC_TIMEOUT_MS);
      }
      STATE.lastSyncAt = new Date().toISOString();
      render("online", "✅ Sincronizzazione completata", 3500);
      window.dispatchEvent(new CustomEvent("hera:offline-sync-complete", { detail: { at: STATE.lastSyncAt } }));
      return true;
    } catch (error) {
      STATE.lastError = String(error?.message || error || "Errore sincronizzazione");
      render("error", "⚠️ Connessione presente, ma alcune modifiche devono ancora sincronizzarsi. Riproverò automaticamente.");
      return false;
    } finally {
      STATE.syncing = false;
    }
  }

  function setOffline() {
    STATE.online = false;
    render("offline", "📴 Modalità offline: puoi continuare a usare i dati già scaricati. Le modifiche Firestore verranno sincronizzate quando torna Internet.");
    window.dispatchEvent(new CustomEvent("hera:offline-mode", { detail: { offline: true } }));
  }

  function setOnline() {
    STATE.online = true;
    window.dispatchEvent(new CustomEvent("hera:offline-mode", { detail: { offline: false } }));
    void syncPendingWrites();
  }

  window.addEventListener("offline", setOffline);
  window.addEventListener("online", setOnline);
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible" && navigator.onLine !== false) void syncPendingWrites();
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      if (navigator.onLine === false) setOffline();
    }, { once: true });
  } else if (navigator.onLine === false) {
    setOffline();
  }

  window.HeraOfflineFirstRuntime = {
    installed: true,
    getState: () => ({ ...STATE }),
    syncNow: syncPendingWrites,
    isOffline: () => navigator.onLine === false
  };
})();

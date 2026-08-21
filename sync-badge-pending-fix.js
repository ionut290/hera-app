(() => {
  "use strict";

  if (window.HeraSyncBadgePendingFix?.installed) return;

  const VERSION = "1.0.0";
  const BADGE_ID = "hera-data-sync-status";
  const OFFLINE_QUEUE_KEY = "heraPendingOfflineMutations";
  const IMPIANTO_QUEUE_KEY = "heraPendingImpiantoActions";
  const SHEET_QUEUE_KEY = "heraPendingSheetExports";
  const WAITING_IMPIANTO_STATUSES = new Set(["pending", "syncing", "syncFailed"]);

  let syncing = false;
  let observer = null;
  let intervalId = null;

  const readArray = (key) => {
    try {
      const value = JSON.parse(localStorage.getItem(key) || "[]");
      return Array.isArray(value) ? value : [];
    } catch (_) {
      return [];
    }
  };

  function countRealPending() {
    const offline = readArray(OFFLINE_QUEUE_KEY)
      .filter((item) => item?.id && item?.type && String(item.status || "pending") !== "synced").length;
    const impianti = readArray(IMPIANTO_QUEUE_KEY)
      .filter((item) => WAITING_IMPIANTO_STATUSES.has(String(item?.status || "pending"))).length;
    const sheets = readArray(SHEET_QUEUE_KEY).length;
    return offline + impianti + sheets;
  }

  function badgeText() {
    const pending = countRealPending();
    if (syncing) return "🔄 Sincronizzazione in corso…";
    if (!navigator.onLine) return pending ? `🟡 Offline · ${pending} da sincronizzare` : "🟡 Offline · dati protetti";
    return pending ? `🟡 ${pending} modifiche da sincronizzare` : "🟢 Sincronizzato";
  }

  function refreshBadge() {
    const badge = document.getElementById(BADGE_ID);
    if (!badge) return;
    const text = badgeText();
    if (badge.textContent !== text) badge.textContent = text;
    badge.disabled = syncing;
    badge.style.cursor = syncing ? "wait" : "pointer";
    badge.title = countRealPending() ? "Tocca per sincronizzare subito le modifiche in attesa" : "Dati sincronizzati";
  }

  async function runWithTimeout(task, timeoutMs = 25000) {
    return Promise.race([
      Promise.resolve().then(task),
      new Promise((_, reject) => setTimeout(() => reject(new Error("Sincronizzazione troppo lenta: riprova")), timeoutMs))
    ]);
  }

  async function forceSync() {
    if (syncing) return { before: countRealPending(), remaining: countRealPending() };
    if (!navigator.onLine) throw new Error("Connessione assente: impossibile sincronizzare adesso");

    const handlers = [
      window.syncPendingOfflineMutations,
      window.syncPendingImpiantoActions,
      window.processPendingSheetExports
    ].filter((handler) => typeof handler === "function");

    if (!handlers.length) throw new Error("Sincronizzazione non disponibile");

    const before = countRealPending();
    syncing = true;
    refreshBadge();
    try {
      await window.HeraDataDurability?.snapshot?.("before-manual-real-sync");
      for (const handler of handlers) {
        await runWithTimeout(() => handler.call(window));
      }
      await new Promise((resolve) => setTimeout(resolve, 150));
      const remaining = countRealPending();
      await window.HeraDataDurability?.snapshot?.("after-manual-real-sync");
      window.dispatchEvent(new CustomEvent("hera:manual-real-sync-complete", { detail: { before, remaining } }));
      return { before, remaining };
    } finally {
      syncing = false;
      refreshBadge();
    }
  }

  function bindBadge() {
    const current = document.getElementById(BADGE_ID);
    if (!current || current.dataset.realPendingFix === "1") return false;

    const badge = current.cloneNode(true);
    badge.dataset.realPendingFix = "1";
    current.replaceWith(badge);
    badge.addEventListener("click", async () => {
      const pending = countRealPending();
      if (!pending) {
        alert("Tutti i dati risultano sincronizzati.");
        refreshBadge();
        return;
      }
      if (!navigator.onLine) {
        alert(`Sei offline. Le ${pending} modifiche restano protette e verranno sincronizzate quando torna la connessione.`);
        return;
      }
      try {
        const result = await forceSync();
        alert(result.remaining > 0
          ? `Sincronizzazione eseguita. Restano ${result.remaining} modifiche realmente in attesa.`
          : `Sincronizzazione completata. ${result.before} modifiche sincronizzate.`);
      } catch (error) {
        alert(`Sincronizzazione non completata: ${String(error?.message || error || "errore sconosciuto")}`);
      }
    });
    refreshBadge();
    return true;
  }

  function installObserver() {
    if (observer || !document.body) return;
    observer = new MutationObserver(() => {
      bindBadge();
      refreshBadge();
    });
    observer.observe(document.body, { childList: true, subtree: true, characterData: true });
  }

  function install() {
    bindBadge();
    installObserver();
    window.addEventListener("storage", refreshBadge);
    window.addEventListener("online", refreshBadge);
    window.addEventListener("offline", refreshBadge);
    window.addEventListener("hera:data-durability-ready", () => { bindBadge(); refreshBadge(); });
    intervalId = window.setInterval(() => { bindBadge(); refreshBadge(); }, 1000);
  }

  window.HeraSyncBadgePendingFix = {
    installed: true,
    version: VERSION,
    countRealPending,
    forceSync,
    refreshBadge,
    stop() {
      if (intervalId) clearInterval(intervalId);
      observer?.disconnect();
      intervalId = null;
      observer = null;
    }
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();

(() => {
  "use strict";

  const VERSION = "1.0.0";
  const DB_NAME = "hera-data-durability-v1";
  const DB_VERSION = 1;
  const SNAPSHOT_STORE = "snapshots";
  const META_STORE = "meta";
  const MAX_BACKUPS = 20;
  const SNAPSHOT_INTERVAL_MS = 15_000;
  const MAX_KEY_BYTES = 1_500_000;
  const MAX_SNAPSHOT_BYTES = 5_000_000;
  const BADGE_ID = "hera-data-sync-status";
  const SECURITY_KEY_PATTERN = /(firebase.*auth|authuser|token|password|passwd|credential|secret|session|salt|private|jwt|access[_-]?key|api[_-]?key)/i;
  const QUEUE_KEY_PATTERN = /(pending|offline|mutation|queue|sync)/i;
  const APP_KEY_PATTERN = /(hera|varga|commess|impiant|squadr|hour|ore|note|segnal|report|fatto|global|mezzi|operator|cantier)/i;

  const state = {
    installed: false,
    ready: false,
    persistentStorage: false,
    lastSnapshotAt: 0,
    lastRestoreAt: 0,
    lastReason: "",
    lastError: "",
    pendingCount: 0,
    snapshotInFlight: null,
    restoreInFlight: null
  };

  function byteSize(value) {
    try {
      return new Blob([String(value ?? "")]).size;
    } catch (_) {
      return String(value ?? "").length * 2;
    }
  }

  function safeJsonParse(value, fallback = null) {
    try {
      return JSON.parse(value);
    } catch (_) {
      return fallback;
    }
  }

  function isSensitiveKey(key) {
    return SECURITY_KEY_PATTERN.test(String(key || ""));
  }

  function isAppDataKey(key) {
    const normalized = String(key || "");
    return !isSensitiveKey(normalized) && (APP_KEY_PATTERN.test(normalized) || QUEUE_KEY_PATTERN.test(normalized));
  }

  function isQueueKey(key) {
    return !isSensitiveKey(key) && QUEUE_KEY_PATTERN.test(String(key || ""));
  }

  function openDb() {
    return new Promise((resolve, reject) => {
      if (!("indexedDB" in window)) {
        reject(new Error("IndexedDB non disponibile"));
        return;
      }
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(SNAPSHOT_STORE)) {
          const store = db.createObjectStore(SNAPSHOT_STORE, { keyPath: "id", autoIncrement: true });
          store.createIndex("createdAt", "createdAt", { unique: false });
        }
        if (!db.objectStoreNames.contains(META_STORE)) {
          db.createObjectStore(META_STORE, { keyPath: "key" });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Apertura IndexedDB fallita"));
      request.onblocked = () => reject(new Error("IndexedDB bloccato da un'altra scheda"));
    });
  }

  async function withStore(storeName, mode, callback) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(storeName, mode);
      const store = tx.objectStore(storeName);
      let result;
      try {
        result = callback(store, tx);
      } catch (error) {
        db.close();
        reject(error);
        return;
      }
      tx.oncomplete = () => {
        db.close();
        resolve(result);
      };
      tx.onerror = () => {
        const error = tx.error || new Error("Transazione IndexedDB fallita");
        db.close();
        reject(error);
      };
      tx.onabort = () => {
        const error = tx.error || new Error("Transazione IndexedDB annullata");
        db.close();
        reject(error);
      };
    });
  }

  async function putMeta(key, value) {
    return withStore(META_STORE, "readwrite", (store) => store.put({ key, value, updatedAt: Date.now() }));
  }

  async function getMeta(key) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(META_STORE, "readonly");
      const request = tx.objectStore(META_STORE).get(key);
      request.onsuccess = () => resolve(request.result?.value);
      request.onerror = () => reject(request.error || new Error("Lettura metadati fallita"));
      tx.oncomplete = () => db.close();
      tx.onerror = () => db.close();
    });
  }

  function collectLocalState() {
    const values = {};
    let totalBytes = 0;
    try {
      for (let i = 0; i < window.localStorage.length; i += 1) {
        const key = window.localStorage.key(i);
        if (!key || !isAppDataKey(key)) continue;
        const value = window.localStorage.getItem(key);
        if (value == null) continue;
        const bytes = byteSize(key) + byteSize(value);
        if (bytes > MAX_KEY_BYTES || totalBytes + bytes > MAX_SNAPSHOT_BYTES) continue;
        values[key] = value;
        totalBytes += bytes;
      }
    } catch (error) {
      state.lastError = String(error?.message || error || "localStorage non leggibile");
    }
    return { values, totalBytes };
  }

  function countPendingValue(key, rawValue) {
    if (!isQueueKey(key) || rawValue == null || rawValue === "") return 0;
    const parsed = safeJsonParse(rawValue, undefined);
    if (Array.isArray(parsed)) return parsed.length;
    if (parsed && typeof parsed === "object") {
      if (Array.isArray(parsed.items)) return parsed.items.length;
      if (Array.isArray(parsed.queue)) return parsed.queue.length;
      if (Array.isArray(parsed.pending)) return parsed.pending.length;
      return Object.keys(parsed).length > 0 ? 1 : 0;
    }
    return String(rawValue).trim() ? 1 : 0;
  }

  function getPendingCount() {
    let count = 0;
    try {
      for (let i = 0; i < window.localStorage.length; i += 1) {
        const key = window.localStorage.key(i);
        if (!key || !isQueueKey(key)) continue;
        count += countPendingValue(key, window.localStorage.getItem(key));
      }
    } catch (_) {}
    state.pendingCount = count;
    updateStatusBadge();
    return count;
  }

  async function pruneBackups() {
    const db = await openDb();
    const rows = await new Promise((resolve, reject) => {
      const tx = db.transaction(SNAPSHOT_STORE, "readonly");
      const request = tx.objectStore(SNAPSHOT_STORE).getAll();
      request.onsuccess = () => resolve(request.result || []);
      request.onerror = () => reject(request.error || new Error("Lettura backup fallita"));
      tx.oncomplete = () => db.close();
      tx.onerror = () => db.close();
    });
    const sorted = rows.sort((a, b) => Number(b.createdAt || 0) - Number(a.createdAt || 0));
    const remove = sorted.slice(MAX_BACKUPS);
    if (!remove.length) return;
    await withStore(SNAPSHOT_STORE, "readwrite", (store) => {
      remove.forEach((row) => store.delete(row.id));
    });
  }

  async function snapshot(reason = "periodic") {
    if (state.snapshotInFlight) return state.snapshotInFlight;
    state.snapshotInFlight = (async () => {
      const { values, totalBytes } = collectLocalState();
      if (!Object.keys(values).length) return null;
      const pendingCount = Object.entries(values).reduce(
        (sum, [key, value]) => sum + countPendingValue(key, value),
        0
      );
      const record = {
        createdAt: Date.now(),
        reason: String(reason || "manual"),
        appVersion: VERSION,
        pendingCount,
        totalBytes,
        values
      };
      await withStore(SNAPSHOT_STORE, "readwrite", (store) => store.add(record));
      state.lastSnapshotAt = record.createdAt;
      state.lastReason = record.reason;
      state.pendingCount = pendingCount;
      await putMeta("lastSnapshot", {
        createdAt: record.createdAt,
        reason: record.reason,
        pendingCount,
        totalBytes
      }).catch(() => null);
      await pruneBackups().catch(() => null);
      updateStatusBadge();
      return record;
    })().catch((error) => {
      state.lastError = String(error?.message || error || "Snapshot fallito");
      updateStatusBadge();
      console.warn("[DATA DURABILITY] Snapshot non riuscito:", error);
      return null;
    }).finally(() => {
      state.snapshotInFlight = null;
    });
    return state.snapshotInFlight;
  }

  async function listBackups() {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(SNAPSHOT_STORE, "readonly");
      const request = tx.objectStore(SNAPSHOT_STORE).getAll();
      request.onsuccess = () => {
        const rows = (request.result || [])
          .sort((a, b) => Number(b.createdAt || 0) - Number(a.createdAt || 0))
          .map(({ values, ...row }) => ({ ...row, keyCount: Object.keys(values || {}).length }));
        resolve(rows);
      };
      request.onerror = () => reject(request.error || new Error("Elenco backup non disponibile"));
      tx.oncomplete = () => db.close();
      tx.onerror = () => db.close();
    });
  }

  async function getBackup(id) {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(SNAPSHOT_STORE, "readonly");
      const request = tx.objectStore(SNAPSHOT_STORE).get(id);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error || new Error("Backup non leggibile"));
      tx.oncomplete = () => db.close();
      tx.onerror = () => db.close();
    });
  }

  async function getLatestBackup() {
    const db = await openDb();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(SNAPSHOT_STORE, "readonly");
      const store = tx.objectStore(SNAPSHOT_STORE);
      const index = store.index("createdAt");
      const request = index.openCursor(null, "prev");
      request.onsuccess = () => resolve(request.result?.value || null);
      request.onerror = () => reject(request.error || new Error("Ultimo backup non leggibile"));
      tx.oncomplete = () => db.close();
      tx.onerror = () => db.close();
    });
  }

  function restoreValues(values, { onlyMissingQueues = false } = {}) {
    let restored = 0;
    Object.entries(values || {}).forEach(([key, value]) => {
      if (!isAppDataKey(key) || isSensitiveKey(key)) return;
      if (onlyMissingQueues && !isQueueKey(key)) return;
      try {
        const current = window.localStorage.getItem(key);
        if (onlyMissingQueues && current != null && current !== "" && current !== "[]" && current !== "{}") return;
        if (byteSize(value) > MAX_KEY_BYTES) return;
        window.localStorage.setItem(key, value);
        restored += 1;
      } catch (error) {
        state.lastError = String(error?.message || error || "Ripristino localStorage fallito");
      }
    });
    return restored;
  }

  async function restoreMissingState() {
    if (state.restoreInFlight) return state.restoreInFlight;
    state.restoreInFlight = (async () => {
      const backup = await getLatestBackup();
      if (!backup?.values) return 0;
      const restored = restoreValues(backup.values, { onlyMissingQueues: true });
      if (restored) {
        state.lastRestoreAt = Date.now();
        await putMeta("lastRestore", {
          restoredAt: state.lastRestoreAt,
          sourceBackupAt: backup.createdAt,
          restoredKeys: restored
        }).catch(() => null);
        window.dispatchEvent(new CustomEvent("hera:data-restored", {
          detail: { restoredKeys: restored, sourceBackupAt: backup.createdAt }
        }));
      }
      getPendingCount();
      return restored;
    })().catch((error) => {
      state.lastError = String(error?.message || error || "Ripristino automatico fallito");
      console.warn("[DATA DURABILITY] Ripristino automatico non riuscito:", error);
      return 0;
    }).finally(() => {
      state.restoreInFlight = null;
    });
    return state.restoreInFlight;
  }

  async function restoreBackup(id, { reload = false } = {}) {
    const backup = await getBackup(id);
    if (!backup?.values) throw new Error("Backup non trovato");
    await snapshot("before-manual-restore");
    const restored = restoreValues(backup.values, { onlyMissingQueues: false });
    state.lastRestoreAt = Date.now();
    await putMeta("lastRestore", {
      restoredAt: state.lastRestoreAt,
      sourceBackupAt: backup.createdAt,
      restoredKeys: restored,
      manual: true
    }).catch(() => null);
    getPendingCount();
    if (reload) window.location.reload();
    return { restoredKeys: restored, sourceBackupAt: backup.createdAt };
  }

  async function prepareForUpdate(reason = "app-update") {
    await snapshot(reason);
    try {
      const handler = window.syncPendingOfflineMutations;
      if (navigator.onLine && typeof handler === "function") {
        await Promise.race([
          Promise.resolve(handler()),
          new Promise((resolve) => window.setTimeout(resolve, 3000))
        ]);
        await snapshot(`${reason}-after-sync`);
      }
    } catch (error) {
      state.lastError = String(error?.message || error || "Sync pre-aggiornamento fallita");
    }
    return { pendingCount: getPendingCount(), snapshotAt: state.lastSnapshotAt };
  }

  function hasPendingOperations() {
    return getPendingCount() > 0;
  }

  function setSyncError(error) {
    state.lastError = error ? String(error?.message || error) : "";
    updateStatusBadge();
  }

  function statusText() {
    if (!navigator.onLine) return state.pendingCount ? `Offline · ${state.pendingCount} da sincronizzare` : "Offline · dati protetti";
    if (state.lastError) return "Sincronizzazione da controllare";
    if (state.pendingCount) return `${state.pendingCount} modifiche da sincronizzare`;
    return "Sincronizzato";
  }

  function statusIcon() {
    if (!navigator.onLine) return "🟡";
    if (state.lastError) return "🔴";
    if (state.pendingCount) return "🟡";
    return "🟢";
  }

  function ensureStatusBadge() {
    if (!document.body || document.getElementById(BADGE_ID)) return;
    const badge = document.createElement("button");
    badge.id = BADGE_ID;
    badge.type = "button";
    badge.setAttribute("aria-live", "polite");
    badge.setAttribute("aria-label", "Stato protezione e sincronizzazione dati");
    badge.title = "Stato protezione dati";
    badge.style.cssText = [
      "position:fixed",
      "right:10px",
      "bottom:max(10px,env(safe-area-inset-bottom))",
      "z-index:2147483000",
      "border:1px solid rgba(15,23,42,.14)",
      "border-radius:999px",
      "background:rgba(255,255,255,.96)",
      "color:#0f172a",
      "box-shadow:0 5px 18px rgba(15,23,42,.14)",
      "padding:6px 9px",
      "font:700 11px/1.1 system-ui,-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif",
      "max-width:min(78vw,260px)",
      "white-space:nowrap",
      "overflow:hidden",
      "text-overflow:ellipsis"
    ].join(";");
    badge.addEventListener("click", async () => {
      const backups = await listBackups().catch(() => []);
      const latest = backups[0];
      const lines = [
        `Stato: ${statusText()}`,
        `Backup locali: ${backups.length}/${MAX_BACKUPS}`,
        latest ? `Ultimo backup: ${new Date(latest.createdAt).toLocaleString("it-IT")}` : "Ultimo backup: non ancora disponibile",
        state.persistentStorage ? "Archiviazione browser: persistente" : "Archiviazione browser: standard"
      ];
      window.alert(lines.join("\n"));
    });
    document.body.appendChild(badge);
  }

  function updateStatusBadge() {
    const badge = document.getElementById(BADGE_ID);
    if (!badge) return;
    badge.textContent = `${statusIcon()} ${statusText()}`;
  }

  async function requestPersistentStorage() {
    try {
      if (!navigator.storage?.persist) return false;
      state.persistentStorage = Boolean(await navigator.storage.persist());
      await putMeta("persistentStorage", state.persistentStorage).catch(() => null);
      return state.persistentStorage;
    } catch (_) {
      return false;
    }
  }

  function installLifecycleGuards() {
    let lastFastSnapshotAt = 0;
    const fastSnapshot = (reason) => {
      const now = Date.now();
      if (now - lastFastSnapshotAt < 800) return;
      lastFastSnapshotAt = now;
      void snapshot(reason);
    };

    window.addEventListener("pagehide", () => fastSnapshot("pagehide"), { capture: true });
    window.addEventListener("offline", () => {
      fastSnapshot("offline");
      getPendingCount();
    });
    window.addEventListener("online", () => {
      getPendingCount();
      window.setTimeout(() => void snapshot("online-after-sync-window"), 2500);
    });
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "hidden") fastSnapshot("visibility-hidden");
      else getPendingCount();
    });
    window.addEventListener("storage", () => getPendingCount());
    window.addEventListener("error", () => fastSnapshot("window-error"));
    window.addEventListener("unhandledrejection", () => fastSnapshot("unhandled-rejection"));

    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.addEventListener("message", (event) => {
        if (event.data?.type !== "HERA_SW_UPDATE_READY") return;
        void prepareForUpdate("service-worker-update-ready").then(() => {
          window.dispatchEvent(new CustomEvent("hera:update-ready", { detail: event.data }));
        });
      });
    }

    window.setInterval(() => {
      getPendingCount();
      void snapshot("periodic");
    }, SNAPSHOT_INTERVAL_MS);
  }

  async function initialize() {
    if (state.installed) return;
    state.installed = true;
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", ensureStatusBadge, { once: true });
    } else {
      ensureStatusBadge();
    }
    try {
      await openDb().then((db) => db.close());
      await restoreMissingState();
      await requestPersistentStorage();
      await snapshot("startup");
      state.ready = true;
      state.lastError = "";
    } catch (error) {
      state.lastError = String(error?.message || error || "Inizializzazione protezione dati fallita");
      console.warn("[DATA DURABILITY] Avvio in modalità degradata:", error);
    }
    getPendingCount();
    installLifecycleGuards();
    updateStatusBadge();
    window.dispatchEvent(new CustomEvent("hera:data-durability-ready", { detail: { ...state } }));
  }

  window.HeraDataDurability = {
    VERSION,
    MAX_BACKUPS,
    DB_NAME,
    ready: () => state.ready,
    snapshot,
    restoreMissingState,
    prepareForUpdate,
    hasPendingOperations,
    getPendingCount,
    listBackups,
    restoreBackup,
    setSyncError,
    getStatus: () => ({ ...state, online: navigator.onLine, text: statusText() }),
    getMeta
  };

  void initialize();
})();

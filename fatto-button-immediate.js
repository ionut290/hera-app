(() => {
  "use strict";

  const DB_NAME = "hera-fatto-sync";
  const DB_VERSION = 1;
  const STORE_NAME = "operations";
  const STALE_SYNC_MS = 2 * 60 * 1000;
  const MAX_ATTEMPTS = 12;
  let processing = false;
  let installed = false;

  const normalize = (value) => String(value ?? "").trim();
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  function makeOperationId(impianto) {
    const key = normalize(
      impianto?.id || impianto?.key || impianto?.idSap || impianto?.sap || impianto?.ID_SAP || impianto?.nome
    ).replace(/[^a-zA-Z0-9_-]+/g, "_");
    const user = normalize(window.auth?.currentUser?.uid || window.auth?.currentUser?.email || "utente")
      .replace(/[^a-zA-Z0-9_-]+/g, "_");
    return `fatto_${key || "impianto"}_${user}_${Date.now()}`;
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
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          const store = db.createObjectStore(STORE_NAME, { keyPath: "operationId" });
          store.createIndex("status", "status", { unique: false });
          store.createIndex("createdAt", "createdAt", { unique: false });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Apertura coda non riuscita"));
    });
  }

  async function withStore(mode, work) {
    const db = await openDb();
    try {
      return await new Promise((resolve, reject) => {
        const transaction = db.transaction(STORE_NAME, mode);
        const store = transaction.objectStore(STORE_NAME);
        let result;
        try {
          result = work(store);
        } catch (error) {
          reject(error);
          return;
        }
        transaction.oncomplete = () => resolve(result);
        transaction.onerror = () => reject(transaction.error || new Error("Errore coda locale"));
        transaction.onabort = () => reject(transaction.error || new Error("Operazione coda annullata"));
      });
    } finally {
      db.close();
    }
  }

  async function putOperation(operation) {
    return withStore("readwrite", (store) => store.put(operation));
  }

  async function getOperations() {
    const db = await openDb();
    try {
      return await new Promise((resolve, reject) => {
        const transaction = db.transaction(STORE_NAME, "readonly");
        const request = transaction.objectStore(STORE_NAME).getAll();
        request.onsuccess = () => resolve(Array.isArray(request.result) ? request.result : []);
        request.onerror = () => reject(request.error || new Error("Lettura coda non riuscita"));
      });
    } finally {
      db.close();
    }
  }

  async function removeOperation(operationId) {
    return withStore("readwrite", (store) => store.delete(operationId));
  }

  function updateSyncIndicator(items = []) {
    const waiting = items.filter((item) => item.status !== "COMPLETED");
    document.documentElement.dataset.fattoSyncPending = String(waiting.length);
    window.dispatchEvent(new CustomEvent("hera:fatto-sync-status", {
      detail: {
        pending: waiting.length,
        syncing: waiting.some((item) => item.status === "SYNCING"),
        failed: waiting.filter((item) => item.status === "FAILED" || item.status === "BLOCKED").length
      }
    }));
  }

  async function refreshIndicator() {
    try {
      updateSyncIndicator(await getOperations());
    } catch (_) {
      // L'indicatore non deve bloccare il flusso FATTO.
    }
  }

  async function enqueue(impianto, metadata = {}) {
    const now = Date.now();
    const operation = {
      operationId: metadata.operationId || makeOperationId(impianto),
      type: "IMPIANTO_FATTO",
      status: "PENDING",
      impianto: JSON.parse(JSON.stringify(impianto || {})),
      commessaId: normalize(metadata.commessaId || window.selectedCommessaId),
      doneAt: metadata.doneAt || new Date(now).toISOString(),
      doneBy: normalize(metadata.doneBy || window.auth?.currentUser?.displayName || window.auth?.currentUser?.email),
      attempts: 0,
      createdAt: now,
      updatedAt: now,
      lastError: ""
    };
    await putOperation(operation);
    await refreshIndicator();
    return operation;
  }

  async function markStatus(operation, status, error = "") {
    const updated = {
      ...operation,
      status,
      updatedAt: Date.now(),
      lastError: normalize(error?.message || error)
    };
    await putOperation(updated);
    return updated;
  }

  async function complete(operation) {
    await removeOperation(operation.operationId);
    await refreshIndicator();
  }

  async function runExistingSync(operation) {
    const impianto = operation.impianto || {};
    const options = {
      source: "resume-persistent-queue",
      operationId: operation.operationId,
      doneAt: operation.doneAt,
      doneBy: operation.doneBy,
      requireFirestoreConfirmation: false,
      reopenWhatsApp: false
    };

    if (typeof window.forceMoveImpiantoToFatti === "function") {
      return window.forceMoveImpiantoToFatti(impianto, options);
    }
    if (typeof window.markImpiantoDone === "function") {
      return window.markImpiantoDone(impianto, options);
    }
    if (typeof window.forceMarkDone === "function") {
      return window.forceMarkDone(impianto, options);
    }
    throw new Error("Funzione di sincronizzazione FATTO non ancora disponibile");
  }

  async function processQueue(reason = "manual") {
    if (processing || !navigator.onLine) return false;
    processing = true;
    document.documentElement.dataset.fattoSyncReason = reason;
    try {
      const items = (await getOperations()).sort((a, b) => a.createdAt - b.createdAt);
      updateSyncIndicator(items);
      for (const original of items) {
        let operation = original;
        if (operation.status === "COMPLETED") {
          await removeOperation(operation.operationId);
          continue;
        }
        if (operation.status === "SYNCING" && Date.now() - Number(operation.updatedAt || 0) < STALE_SYNC_MS) continue;
        if (Number(operation.attempts || 0) >= MAX_ATTEMPTS) {
          await markStatus(operation, "BLOCKED", operation.lastError || "Numero massimo di tentativi raggiunto");
          continue;
        }
        operation = await markStatus({ ...operation, attempts: Number(operation.attempts || 0) + 1 }, "SYNCING");
        try {
          await runExistingSync(operation);
          await complete(operation);
        } catch (error) {
          await markStatus(operation, "FAILED", error);
          if (/non ancora disponibile/i.test(normalize(error?.message))) break;
          await sleep(Math.min(800 * operation.attempts, 4000));
        }
      }
      await refreshIndicator();
      return true;
    } finally {
      processing = false;
    }
  }

  function isNativeAndroid() {
    try {
      return window.Capacitor?.getPlatform?.() === "android";
    } catch (_) {
      return false;
    }
  }

  function getNativeWhatsAppPlugin() {
    return window.Capacitor?.Plugins?.HeraWhatsApp || null;
  }

  function extractWhatsAppUrl(args) {
    for (const value of args) {
      if (typeof value === "string" && /(whatsapp:|wa\.me|api\.whatsapp\.com)/i.test(value)) return value;
      if (value && typeof value === "object") {
        const candidate = value.url || value.href || value.link || value.whatsappUrl;
        if (typeof candidate === "string" && candidate) return candidate;
      }
    }
    return "";
  }

  function installOpenWhatsAppWrapper() {
    const original = window.openWhatsApp;
    if (typeof original !== "function" || original.__heraNativeWhatsAppWrapped) return false;

    const wrapped = function heraOpenInstalledWhatsApp(...args) {
      const plugin = getNativeWhatsAppPlugin();
      if (isNativeAndroid() && plugin?.open) {
        const url = extractWhatsAppUrl(args);
        plugin.open({ url }).catch((error) => {
          console.error("Apertura WhatsApp nativa fallita:", error);
          window.dispatchEvent(new CustomEvent("hera:whatsapp-error", { detail: { message: normalize(error?.message || error) } }));
        });
        return true;
      }
      return original.apply(this, args);
    };
    wrapped.__heraNativeWhatsAppWrapped = true;
    wrapped.__original = original;
    window.openWhatsApp = wrapped;
    return true;
  }

  function installFattoFlowWrapper() {
    const original = window.handleImpiantoWhatsAppClick;
    if (typeof original !== "function" || original.__heraPersistentQueueWrapped) return false;

    const wrapped = async function heraPersistentFattoFlow(impianto, ...args) {
      let operation;
      try {
        operation = await enqueue(impianto, {
          commessaId: window.selectedCommessaId,
          doneAt: new Date().toISOString()
        });
      } catch (error) {
        console.error("Impossibile mettere FATTO in sicurezza:", error);
        throw error;
      }

      try {
        const result = await original.call(this, impianto, ...args);
        if (result === true) await complete(operation);
        else await markStatus(operation, "FAILED", "Flusso FATTO non completato");
        return result;
      } catch (error) {
        await markStatus(operation, "FAILED", error);
        throw error;
      }
    };
    wrapped.__heraPersistentQueueWrapped = true;
    wrapped.__original = original;
    window.handleImpiantoWhatsAppClick = wrapped;
    return true;
  }

  function installWrappers() {
    installOpenWhatsAppWrapper();
    installFattoFlowWrapper();
  }

  function requestResume(reason) {
    installWrappers();
    window.setTimeout(() => processQueue(reason), 80);
  }

  function installLifecycle() {
    if (installed) return;
    installed = true;

    window.addEventListener("online", () => requestResume("online"));
    window.addEventListener("pageshow", () => requestResume("pageshow"));
    window.addEventListener("focus", () => requestResume("focus"));
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") requestResume("visible");
    });
    window.addEventListener("hera:auth-ready", () => requestResume("auth-ready"));
    window.addEventListener("hera:data-ready", () => requestResume("data-ready"));
    window.addEventListener("hera:native-resume", () => requestResume("native-resume"));

    const appPlugin = window.Capacitor?.Plugins?.App;
    if (appPlugin?.addListener) {
      appPlugin.addListener("appStateChange", ({ isActive }) => {
        if (isActive) requestResume("capacitor-active");
      }).catch?.(() => {});
    }

    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      installWrappers();
      if (attempts >= 120 || (
        window.openWhatsApp?.__heraNativeWhatsAppWrapped
        && window.handleImpiantoWhatsAppClick?.__heraPersistentQueueWrapped
      )) window.clearInterval(timer);
    }, 250);

    refreshIndicator();
    requestResume("startup");
  }

  window.HeraFattoSync = Object.freeze({
    enqueue,
    processQueue,
    refreshIndicator,
    getOperations
  });

  installLifecycle();
})();

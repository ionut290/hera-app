(() => {
  "use strict";

  const DB_NAME = "hera-fatto-sync";
  const STORE = "operations";
  const MAX_ATTEMPTS = 12;
  const STALE_MS = 120000;
  let processing = false;

  const text = (value) => String(value ?? "").trim();
  const clone = (value) => JSON.parse(JSON.stringify(value || {}));

  function openDb() {
    return new Promise((resolve, reject) => {
      if (!("indexedDB" in window)) return reject(new Error("IndexedDB non disponibile"));
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => {
        if (!request.result.objectStoreNames.contains(STORE)) {
          request.result.createObjectStore(STORE, { keyPath: "operationId" });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Coda FATTO non disponibile"));
    });
  }

  async function transact(mode, callback) {
    const db = await openDb();
    try {
      return await new Promise((resolve, reject) => {
        const transaction = db.transaction(STORE, mode);
        const store = transaction.objectStore(STORE);
        const result = callback(store);
        transaction.oncomplete = () => resolve(result);
        transaction.onerror = () => reject(transaction.error || new Error("Errore coda FATTO"));
        transaction.onabort = () => reject(transaction.error || new Error("Coda FATTO annullata"));
      });
    } finally {
      db.close();
    }
  }

  const put = (operation) => transact("readwrite", (store) => store.put(operation));
  const remove = (operationId) => transact("readwrite", (store) => store.delete(operationId));

  async function list() {
    const db = await openDb();
    try {
      return await new Promise((resolve, reject) => {
        const request = db.transaction(STORE, "readonly").objectStore(STORE).getAll();
        request.onsuccess = () => resolve(Array.isArray(request.result) ? request.result : []);
        request.onerror = () => reject(request.error || new Error("Lettura coda FATTO fallita"));
      });
    } finally {
      db.close();
    }
  }

  function publishStatus(items) {
    const pending = items.filter((item) => item.status !== "COMPLETED");
    document.documentElement.dataset.fattoSyncPending = String(pending.length);
    window.dispatchEvent(new CustomEvent("hera:fatto-sync-status", {
      detail: {
        pending: pending.length,
        syncing: pending.some((item) => item.status === "SYNCING"),
        failed: pending.filter((item) => ["FAILED", "BLOCKED"].includes(item.status)).length
      }
    }));
  }

  async function refreshStatus() {
    try { publishStatus(await list()); } catch (_) {}
  }

  function createOperationId(impianto) {
    const plant = text(impianto?.id || impianto?.idSap || impianto?.sap || impianto?.nome || "impianto")
      .replace(/[^a-zA-Z0-9_-]+/g, "_");
    const user = text(window.auth?.currentUser?.uid || window.auth?.currentUser?.email || "utente")
      .replace(/[^a-zA-Z0-9_-]+/g, "_");
    return `fatto_${plant}_${user}_${Date.now()}`;
  }

  async function enqueue(impianto, metadata = {}) {
    const now = Date.now();
    const operation = {
      operationId: metadata.operationId || createOperationId(impianto),
      type: "IMPIANTO_FATTO",
      status: "PENDING",
      impianto: clone(impianto),
      commessaId: text(metadata.commessaId || window.selectedCommessaId),
      doneAt: metadata.doneAt || new Date(now).toISOString(),
      doneBy: text(metadata.doneBy || window.auth?.currentUser?.displayName || window.auth?.currentUser?.email),
      attempts: 0,
      createdAt: now,
      updatedAt: now,
      lastError: ""
    };
    await put(operation);
    await refreshStatus();
    return operation;
  }

  async function setStatus(operation, status, error = "") {
    const updated = {
      ...operation,
      status,
      updatedAt: Date.now(),
      lastError: text(error?.message || error)
    };
    await put(updated);
    return updated;
  }

  async function syncOperation(operation) {
    const options = {
      source: "resume-persistent-queue",
      operationId: operation.operationId,
      doneAt: operation.doneAt,
      doneBy: operation.doneBy,
      requireFirestoreConfirmation: false,
      reopenWhatsApp: false
    };
    if (typeof window.forceMoveImpiantoToFatti === "function") {
      return window.forceMoveImpiantoToFatti(operation.impianto, options);
    }
    if (typeof window.markImpiantoDone === "function") {
      return window.markImpiantoDone(operation.impianto, options);
    }
    throw new Error("Funzione FATTO non ancora pronta");
  }

  async function processQueue(reason = "manual") {
    if (processing || !navigator.onLine) return false;
    processing = true;
    document.documentElement.dataset.fattoSyncReason = reason;
    try {
      const items = (await list()).sort((a, b) => a.createdAt - b.createdAt);
      publishStatus(items);
      for (let operation of items) {
        if (operation.status === "SYNCING" && Date.now() - Number(operation.updatedAt || 0) < STALE_MS) continue;
        if (Number(operation.attempts || 0) >= MAX_ATTEMPTS) {
          await setStatus(operation, "BLOCKED", operation.lastError || "Troppi tentativi");
          continue;
        }
        operation = await setStatus({ ...operation, attempts: Number(operation.attempts || 0) + 1 }, "SYNCING");
        try {
          await syncOperation(operation);
          await remove(operation.operationId);
        } catch (error) {
          await setStatus(operation, "FAILED", error);
          if (/non ancora pronta/i.test(text(error?.message))) break;
        }
      }
      await refreshStatus();
      return true;
    } finally {
      processing = false;
    }
  }

  function isAndroidNative() {
    try { return window.Capacitor?.getPlatform?.() === "android"; } catch (_) { return false; }
  }

  function buildNativeWhatsAppUrl(args) {
    for (const value of args) {
      if (typeof value === "string" && /(whatsapp:|wa\.me|api\.whatsapp\.com)/i.test(value)) return value;
      if (value && typeof value === "object") {
        const candidate = value.url || value.href || value.link || value.whatsappUrl;
        if (typeof candidate === "string" && candidate) return candidate;
      }
    }
    const message = args.find((value) => typeof value === "string" && text(value));
    return message ? `https://wa.me/?text=${encodeURIComponent(message)}` : "https://wa.me/";
  }

  function installWhatsAppWrapper() {
    const original = window.openWhatsApp;
    if (typeof original !== "function" || original.__heraNativeWrapped) return;
    const wrapped = function (...args) {
      const plugin = window.Capacitor?.Plugins?.HeraWhatsApp;
      if (isAndroidNative() && plugin?.open) {
        plugin.open({ url: buildNativeWhatsAppUrl(args) }).catch((error) => {
          const message = text(error?.message || error || "WhatsApp non è installato sul dispositivo.");
          window.dispatchEvent(new CustomEvent("hera:whatsapp-error", { detail: { message } }));
          if (typeof window.alert === "function") window.alert(message);
        });
        return true;
      }
      return original.apply(this, args);
    };
    wrapped.__heraNativeWrapped = true;
    wrapped.__original = original;
    window.openWhatsApp = wrapped;
  }

  function installFattoWrapper() {
    const original = window.handleImpiantoWhatsAppClick;
    if (typeof original !== "function" || original.__heraQueueWrapped) return;
    const wrapped = async function (impianto, ...args) {
      const operation = await enqueue(impianto, { commessaId: window.selectedCommessaId });
      try {
        const result = await original.call(this, impianto, ...args);
        if (result === true) await remove(operation.operationId);
        else await setStatus(operation, "FAILED", "Flusso FATTO non completato");
        await refreshStatus();
        return result;
      } catch (error) {
        await setStatus(operation, "FAILED", error);
        await refreshStatus();
        throw error;
      }
    };
    wrapped.__heraQueueWrapped = true;
    wrapped.__original = original;
    window.handleImpiantoWhatsAppClick = wrapped;
  }

  function resume(reason) {
    installWhatsAppWrapper();
    installFattoWrapper();
    window.setTimeout(() => processQueue(reason), 100);
  }

  window.addEventListener("online", () => resume("online"));
  window.addEventListener("pageshow", () => resume("pageshow"));
  window.addEventListener("focus", () => resume("focus"));
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") resume("visible");
  });
  window.addEventListener("hera:auth-ready", () => resume("auth-ready"));
  window.addEventListener("hera:data-ready", () => resume("data-ready"));
  window.addEventListener("hera:native-resume", () => resume("native-resume"));

  let wrapperAttempts = 0;
  const wrapperTimer = window.setInterval(() => {
    wrapperAttempts += 1;
    installWhatsAppWrapper();
    installFattoWrapper();
    if (wrapperAttempts >= 120 || (window.openWhatsApp?.__heraNativeWrapped && window.handleImpiantoWhatsAppClick?.__heraQueueWrapped)) {
      window.clearInterval(wrapperTimer);
    }
  }, 250);

  window.HeraFattoSync = Object.freeze({ enqueue, processQueue, refreshStatus, list });
  refreshStatus();
  resume("startup");
})();

(() => {
  "use strict";
  if (document.querySelector('script[data-password-access-manager="true"]')) return;
  const script = document.createElement("script");
  script.src = "password-access-manager.js?v=20260727a";
  script.defer = true;
  script.dataset.passwordAccessManager = "true";
  document.head.appendChild(script);
})();

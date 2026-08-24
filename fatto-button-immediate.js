(() => {
  "use strict";

  const DB_NAME = "hera-fatto-sync";
  const STORE = "operations";
  const FALLBACK_STORE_KEY = "heraFattoSyncOperationsFallbackV1";
  const MAX_ATTEMPTS = 12;
  const STALE_MS = 120000;
  const YELLOW = "#f4c542";
  const YELLOW_BORDER = "#c99700";
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

  function readFallbackOperations() {
    try {
      if (!window.localStorage) return [];
      const raw = window.localStorage.getItem(FALLBACK_STORE_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed) ? parsed.filter((entry) => entry?.operationId) : [];
    } catch (error) {
      console.warn("[Cassaforte FATTO] ripiego locale non leggibile", error);
      return [];
    }
  }

  function writeFallbackOperations(items) {
    try {
      if (!window.localStorage) return false;
      window.localStorage.setItem(FALLBACK_STORE_KEY, JSON.stringify(items || []));
      return true;
    } catch (error) {
      console.warn("[Cassaforte FATTO] ripiego locale non scrivibile", error);
      return false;
    }
  }

  function upsertFallbackOperation(operation) {
    const operations = readFallbackOperations();
    const index = operations.findIndex((entry) => entry.operationId === operation.operationId);
    if (index >= 0) operations[index] = clone(operation);
    else operations.push(clone(operation));
    return writeFallbackOperations(operations);
  }

  function removeFallbackOperation(operationId) {
    const operations = readFallbackOperations();
    const next = operations.filter((entry) => entry.operationId !== operationId);
    return next.length === operations.length || writeFallbackOperations(next);
  }

  async function put(operation) {
    try {
      await transact("readwrite", (store) => store.put(operation));
      removeFallbackOperation(operation.operationId);
      return "indexeddb";
    } catch (error) {
      if (upsertFallbackOperation(operation)) {
        console.warn("[Cassaforte FATTO] IndexedDB non disponibile: operazione salvata nel ripiego locale", error);
        return "localstorage";
      }
      throw error;
    }
  }

  async function remove(operationId) {
    let removedFromIndexedDb = false;
    try {
      await transact("readwrite", (store) => store.delete(operationId));
      removedFromIndexedDb = true;
    } catch (error) {
      console.warn("[Cassaforte FATTO] rimozione IndexedDB rinviata", error);
    }
    const removedFromFallback = removeFallbackOperation(operationId);
    return removedFromIndexedDb || removedFromFallback;
  }

  async function list() {
    let indexedOperations = [];
    try {
      const db = await openDb();
      try {
        indexedOperations = await new Promise((resolve, reject) => {
          const request = db.transaction(STORE, "readonly").objectStore(STORE).getAll();
          request.onsuccess = () => resolve(Array.isArray(request.result) ? request.result : []);
          request.onerror = () => reject(request.error || new Error("Lettura coda FATTO fallita"));
        });
      } finally {
        db.close();
      }
    } catch (error) {
      console.warn("[Cassaforte FATTO] lettura IndexedDB non disponibile: uso il ripiego locale", error);
    }

    const merged = new Map();
    [...indexedOperations, ...readFallbackOperations()].forEach((operation) => {
      if (!operation?.operationId) return;
      const previous = merged.get(operation.operationId);
      if (!previous || Number(operation.updatedAt || 0) >= Number(previous.updatedAt || 0)) {
        merged.set(operation.operationId, operation);
      }
    });
    return [...merged.values()];
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

  async function enqueueSafely(impianto, metadata = {}) {
    try {
      return await enqueue(impianto, metadata);
    } catch (error) {
      console.error("[Cassaforte FATTO] salvataggio locale non disponibile; il flusso operativo continua", error);
      try {
        window.dispatchEvent(new CustomEvent("hera:fatto-vault-error", {
          detail: { message: text(error?.message || error), impianto: clone(impianto) }
        }));
      } catch (_) {}
      return null;
    }
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

  function formatDoneLabel(doneAt) {
    const date = new Date(doneAt);
    if (Number.isNaN(date.getTime())) return "";
    return new Intl.DateTimeFormat("it-IT", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit"
    }).format(date);
  }

  function findPressedFattoButton() {
    const active = document.activeElement;
    if (!(active instanceof HTMLElement)) return null;
    const button = active.closest("button, [role='button'], input[type='button'], input[type='submit']");
    if (!(button instanceof HTMLElement)) return null;
    const label = text(button.textContent || button.getAttribute("value") || button.getAttribute("aria-label"));
    return /fatto|whazzup|whatsapp/i.test(label) ? button : null;
  }

  function applyPermanentYellowFeedback(button, doneAt) {
    if (!(button instanceof HTMLElement)) return;
    const label = formatDoneLabel(doneAt);
    button.dataset.fattoImmediate = "true";
    button.dataset.fattoDoneAt = doneAt;
    button.disabled = true;
    button.setAttribute("aria-disabled", "true");
    button.style.setProperty("background", YELLOW, "important");
    button.style.setProperty("background-color", YELLOW, "important");
    button.style.setProperty("border-color", YELLOW_BORDER, "important");
    button.style.setProperty("color", "#1d1d1d", "important");
    button.style.setProperty("opacity", "1", "important");
    button.style.setProperty("pointer-events", "none", "important");
    button.style.setProperty("cursor", "default", "important");

    let dateNode = button.previousElementSibling;
    if (!(dateNode instanceof HTMLElement) || dateNode.dataset.fattoImmediateDate !== "true") {
      dateNode = document.createElement("div");
      dateNode.dataset.fattoImmediateDate = "true";
      dateNode.style.fontWeight = "700";
      dateNode.style.fontSize = "0.82rem";
      dateNode.style.marginBottom = "4px";
      dateNode.style.textAlign = "center";
      button.parentNode?.insertBefore(dateNode, button);
    }
    dateNode.textContent = label;

    if ("value" in button && /^(INPUT|BUTTON)$/i.test(button.tagName)) {
      if (button.tagName === "INPUT") button.value = "FATTO";
      else button.textContent = "FATTO";
    } else {
      button.textContent = "FATTO";
    }
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

    const impianto = args[0];
    const options = args[1] && typeof args[1] === "object" ? args[1] : {};
    if (impianto && typeof impianto === "object" && typeof window.buildImpiantoWhatsAppPayload === "function") {
      try {
        const payload = window.buildImpiantoWhatsAppPayload(impianto, options);
        if (typeof payload?.appUrl === "string" && payload.appUrl) return payload.appUrl;
        if (typeof payload?.webUrl === "string" && payload.webUrl) return payload.webUrl;
        if (typeof payload?.message === "string" && text(payload.message)) {
          return `whatsapp://send?text=${encodeURIComponent(payload.message)}`;
        }
      } catch (error) {
        console.error("Errore preparazione messaggio WhatsApp nativo:", error);
      }
    }

    const message = args.find((value) => typeof value === "string" && text(value));
    return message ? `whatsapp://send?text=${encodeURIComponent(message)}` : "";
  }

  function installWhatsAppWrapper() {
    const original = window.openWhatsApp;
    if (typeof original !== "function" || original.__heraNativeWrapped) return;
    const wrapped = function (...args) {
      const plugin = window.Capacitor?.Plugins?.HeraWhatsApp;
      if (isAndroidNative() && plugin?.open) {
        const nativeUrl = buildNativeWhatsAppUrl(args);
        if (!nativeUrl) return original.apply(this, args);
        plugin.open({ url: nativeUrl }).catch((error) => {
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
      const pressedButton = findPressedFattoButton();
      const doneAt = new Date().toISOString();
      applyPermanentYellowFeedback(pressedButton, doneAt);

      const operation = await enqueueSafely(impianto, {
        commessaId: window.selectedCommessaId,
        doneAt
      });
      try {
        const result = await original.call(this, impianto, ...args);
        if (operation) {
          try {
            if (result === true) await remove(operation.operationId);
            else await setStatus(operation, "FAILED", "Flusso FATTO non completato");
          } catch (vaultError) {
            console.error("[Cassaforte FATTO] aggiornamento stato coda rinviato", vaultError);
          }
        }
        await refreshStatus();
        return result;
      } catch (error) {
        if (operation) {
          try { await setStatus(operation, "FAILED", error); }
          catch (vaultError) { console.error("[Cassaforte FATTO] registrazione errore rinviata", vaultError); }
        }
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

  window.HeraFattoSync = Object.freeze({ enqueue, enqueueSafely, processQueue, refreshStatus, list });
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
(function installFirestoreStartupReadOptimizer() {
  "use strict";

  if (window.HeraFirestoreStartupReadOptimizer?.installed) return;
  if (typeof db === "undefined" || !db || typeof firebase === "undefined" || !firebase.firestore) return;

  const state = {
    installed: true,
    version: "1.0.0",
    commesseInitialReadsCoalesced: 0,
    posDuplicateStartsSkipped: 0
  };

  function queryPath(query) {
    if (typeof query?.path === "string") return query.path;
    const path = query?._query?.path || query?._delegate?._query?.path;
    if (path && typeof path.canonicalString === "function") return path.canonicalString();
    return path ? String(path) : "";
  }

  function observerFromArgs(args) {
    let index = 0;
    let options = null;
    const first = args[0];
    const firstIsObserver = first && typeof first === "object"
      && (typeof first.next === "function" || typeof first.error === "function" || typeof first.complete === "function");
    if (first && typeof first === "object" && !firstIsObserver) {
      options = first;
      index = 1;
    }
    const candidate = args[index];
    if (candidate && typeof candidate === "object") {
      return {
        options,
        next: typeof candidate.next === "function" ? candidate.next.bind(candidate) : null,
        error: typeof candidate.error === "function" ? candidate.error.bind(candidate) : null,
        complete: typeof candidate.complete === "function" ? candidate.complete.bind(candidate) : null
      };
    }
    return {
      options,
      next: typeof args[index] === "function" ? args[index] : null,
      error: typeof args[index + 1] === "function" ? args[index + 1] : null,
      complete: typeof args[index + 2] === "function" ? args[index + 2] : null
    };
  }

  const QueryPrototype = firebase.firestore.Query?.prototype;
  const sharedCommesse = new WeakMap();

  if (QueryPrototype && !QueryPrototype.__heraStartupDedupOriginalOnSnapshot) {
    const sourceOnSnapshot = QueryPrototype.onSnapshot;
    Object.defineProperty(QueryPrototype, "__heraStartupDedupOriginalOnSnapshot", {
      value: sourceOnSnapshot,
      configurable: false,
      enumerable: false,
      writable: false
    });

    QueryPrototype.onSnapshot = function startupDedupOnSnapshot(...args) {
      if (queryPath(this) !== "commesse") {
        return sourceOnSnapshot.apply(this, args);
      }

      const subscriber = observerFromArgs(args);
      let entry = sharedCommesse.get(this);
      if (!entry) {
        entry = {
          query: this,
          subscribers: new Set(),
          backendUnsubscribe: null,
          closeTimer: null,
          options: subscriber.options
        };
        sharedCommesse.set(this, entry);

        const backendObserver = {
          next(snapshot) {
            [...entry.subscribers].forEach((item) => {
              try { item.next?.(snapshot); } catch (error) { console.error("Errore callback commesse condivisa:", error); }
            });
          },
          error(error) {
            [...entry.subscribers].forEach((item) => {
              try { item.error?.(error); } catch (callbackError) { console.error("Errore callback errore commesse:", callbackError); }
            });
            entry.subscribers.clear();
            sharedCommesse.delete(entry.query);
          },
          complete() {
            [...entry.subscribers].forEach((item) => {
              try { item.complete?.(); } catch (error) { console.error("Errore callback completamento commesse:", error); }
            });
            entry.subscribers.clear();
            sharedCommesse.delete(entry.query);
          }
        };

        entry.backendUnsubscribe = entry.options
          ? sourceOnSnapshot.call(this, entry.options, backendObserver)
          : sourceOnSnapshot.call(this, backendObserver);
      } else if (entry.closeTimer) {
        clearTimeout(entry.closeTimer);
        entry.closeTimer = null;
      }

      entry.subscribers.add(subscriber);
      let active = true;
      return () => {
        if (!active) return;
        active = false;
        entry.subscribers.delete(subscriber);
        if (entry.subscribers.size || entry.closeTimer) return;
        entry.closeTimer = setTimeout(() => {
          entry.closeTimer = null;
          if (entry.subscribers.size) return;
          try { entry.backendUnsubscribe?.(); } catch (error) { console.warn("Errore chiusura listener commesse condiviso:", error); }
          sharedCommesse.delete(entry.query);
        }, 1500);
      };
    };
  }

  if (typeof runFirestoreGetWithRetry === "function" && !runFirestoreGetWithRetry.__heraStartupDedupWrapped) {
    const sourceRunFirestoreGetWithRetry = runFirestoreGetWithRetry;
    const optimizedRunFirestoreGetWithRetry = function optimizedRunFirestoreGetWithRetry(query, options = {}) {
      if (queryPath(query) !== "commesse" || String(options?.label || "") !== "LOAD COMMESSE") {
        return sourceRunFirestoreGetWithRetry(query, options);
      }

      return new Promise((resolve, reject) => {
        let settled = false;
        let unsubscribe = () => {};
        const timeoutMs = Math.max(3000, Number(options?.timeoutMs) || 9000);
        const timer = setTimeout(() => {
          if (settled) return;
          settled = true;
          try { unsubscribe(); } catch (_) {}
          sourceRunFirestoreGetWithRetry(query, options).then(resolve, reject);
        }, timeoutMs);

        try {
          unsubscribe = query.onSnapshot((snapshot) => {
            if (settled) return;
            settled = true;
            clearTimeout(timer);
            state.commesseInitialReadsCoalesced += Number(snapshot?.size || snapshot?.docs?.length || 0);
            resolve(snapshot);
            queueMicrotask(() => {
              try { unsubscribe(); } catch (_) {}
            });
          }, (error) => {
            if (settled) return;
            settled = true;
            clearTimeout(timer);
            try { unsubscribe(); } catch (_) {}
            sourceRunFirestoreGetWithRetry(query, options).then(resolve, reject).catch(reject);
          });
        } catch (error) {
          clearTimeout(timer);
          sourceRunFirestoreGetWithRetry(query, options).then(resolve, reject).catch(reject);
        }
      });
    };
    Object.defineProperty(optimizedRunFirestoreGetWithRetry, "__heraStartupDedupWrapped", { value: true });
    runFirestoreGetWithRetry = optimizedRunFirestoreGetWithRetry;
  }

  if (typeof subscribePosDocuments === "function" && typeof stopPosDocumentsSubscription === "function") {
    const sourceSubscribePosDocuments = subscribePosDocuments;
    const sourceStopPosDocumentsSubscription = stopPosDocumentsSubscription;
    let activePosKey = "";
    let internalRestart = false;

    subscribePosDocuments = function subscribePosDocumentsDeduplicated() {
      const key = `${String(currentUser?.uid || "anonymous")}|${canManageData() ? "admin" : "user"}`;
      if (activePosKey === key && typeof unsubscribePosDocuments === "function") {
        state.posDuplicateStartsSkipped += 1;
        renderPosDocuments?.();
        return unsubscribePosDocuments;
      }

      internalRestart = true;
      try {
        const result = sourceSubscribePosDocuments();
        activePosKey = typeof unsubscribePosDocuments === "function" ? key : "";
        return result;
      } finally {
        internalRestart = false;
      }
    };

    stopPosDocumentsSubscription = function stopPosDocumentsSubscriptionDeduplicated() {
      const result = sourceStopPosDocumentsSubscription();
      if (!internalRestart) activePosKey = "";
      return result;
    };
  }

  window.HeraFirestoreStartupReadOptimizer = {
    installed: true,
    version: state.version,
    getState: () => ({ ...state })
  };
})();

(function () {
  "use strict";
  const TECHNICAL_COLUMNS = ["RECORD_ID", "UPDATED_AT", "UPDATED_BY", "SYNC_VERSION", "SYNC_SOURCE", "ROW_STATUS"];
  const SHEETS = ["PERSONALE", "MEZZI", "COMMESSE_PERSONALE", "COMMESSE_MEZZI", "LOG_SINCRONIZZAZIONE"];
  const CONFIG_FIELD = "registryGoogleSheetLinks";
  const AUTO_SYNC_MIN_INTERVAL_MS = 30 * 60 * 1000;
  const configs = new Map();
  const syncInFlight = new Map();
  const autoSyncChecked = new Set();
  let configDocumentPromise = null;
  let configDocumentCache = null;
  const clean = (value) => String(value == null ? "" : value).trim();
  const configDocRef = () => db.collection("appConfig").doc("driveBridge");

  function getSharedDriveBridgeConfigState() {
    const root = typeof window !== "undefined" ? window : globalThis;
    const existing = root.HeraDriveBridgeConfigShared;
    if (existing && existing.firstSnapshot && typeof existing.publish === "function") return existing;

    let resolveFirstSnapshot = null;
    const firstSnapshot = new Promise((resolve) => {
      resolveFirstSnapshot = resolve;
    });
    const state = {
      hasSnapshot: false,
      data: null,
      firstSnapshot,
      publish(data) {
        state.data = data && typeof data === "object" ? data : {};
        if (state.hasSnapshot) return;
        state.hasSnapshot = true;
        resolveFirstSnapshot(state.data);
      }
    };
    root.HeraDriveBridgeConfigShared = state;
    return state;
  }

  const sharedDriveBridgeConfig = getSharedDriveBridgeConfigState();

  async function waitForSharedDriveBridgeConfig(timeoutMs = 900) {
    if (sharedDriveBridgeConfig.hasSnapshot) return sharedDriveBridgeConfig.data || {};
    let timer = null;
    try {
      return await Promise.race([
        sharedDriveBridgeConfig.firstSnapshot,
        new Promise((resolve) => {
          timer = setTimeout(() => resolve(null), timeoutMs);
        })
      ]);
    } finally {
      if (timer) clearTimeout(timer);
    }
  }
  const timestamp = (value) => value?.toDate instanceof Function ? value.toDate().toISOString() : clean(value) || new Date(0).toISOString();
  const autoSyncKey = (type) => `hera_registry_sheet_auto_sync_${type}`;

  function friendlyError(error) {
    const code = clean(error?.code).toLowerCase();
    const message = clean(error?.message);
    if (code.includes("permission-denied") || /missing or insufficient permissions/i.test(message)) {
      return "Non è possibile leggere o modificare la configurazione del Foglio Google. Verifica di essere collegato come amministratore.";
    }
    if (code.includes("unauthenticated") || /sessione scaduta/i.test(message)) {
      return "Sessione non disponibile. Accedi nuovamente all’app.";
    }
    return message || "Operazione sul Foglio Google non riuscita.";
  }

  async function readConfigDocument(force = false) {
    if (!force && sharedDriveBridgeConfig.hasSnapshot) {
      configDocumentCache = sharedDriveBridgeConfig.data || {};
      return configDocumentCache;
    }
    if (!force && configDocumentCache) return configDocumentCache;
    if (!force && configDocumentPromise) return configDocumentPromise;

    if (!force) {
      configDocumentPromise = waitForSharedDriveBridgeConfig()
        .then((sharedData) => {
          if (sharedData !== null) {
            configDocumentCache = sharedData || {};
            return configDocumentCache;
          }
          return configDocRef().get().then((snap) => {
            configDocumentCache = snap.exists ? snap.data() || {} : {};
            return configDocumentCache;
          });
        })
        .finally(() => {
          configDocumentPromise = null;
        });
      return configDocumentPromise;
    }

    configDocumentPromise = configDocRef().get()
      .then((snap) => {
        configDocumentCache = snap.exists ? snap.data() || {} : {};
        return configDocumentCache;
      })
      .finally(() => {
        configDocumentPromise = null;
      });
    return configDocumentPromise;
  }

  async function load(type, force = false) {
    if (!force && configs.has(type)) return configs.get(type);
    const data = await readConfigDocument(force);
    const links = data && typeof data[CONFIG_FIELD] === "object" ? data[CONFIG_FIELD] : {};
    const config = links && typeof links[type] === "object" ? links[type] : {};
    configs.set(type, config);
    return config;
  }

  async function save(type, patch) {
    const existing = await load(type);
    const next = { ...existing, ...patch };
    await configDocRef().set({ [CONFIG_FIELD]: { [type]: next } }, { merge: true });
    const previousLinks = configDocumentCache && typeof configDocumentCache[CONFIG_FIELD] === "object"
      ? configDocumentCache[CONFIG_FIELD]
      : {};
    configDocumentCache = {
      ...(configDocumentCache || {}),
      [CONFIG_FIELD]: { ...previousLinks, [type]: next }
    };
    configs.set(type, next);
    return next;
  }

  async function remove(type) {
    try {
      await configDocRef().update({ [`${CONFIG_FIELD}.${type}`]: firebase.firestore.FieldValue.delete() });
    } catch (error) {
      if (clean(error?.code).toLowerCase() !== "not-found") throw error;
    }
    configs.delete(type);
    if (configDocumentCache && typeof configDocumentCache[CONFIG_FIELD] === "object") {
      const links = { ...configDocumentCache[CONFIG_FIELD] };
      delete links[type];
      configDocumentCache = { ...configDocumentCache, [CONFIG_FIELD]: links };
    }
    autoSyncChecked.delete(type);
    try { sessionStorage.removeItem(autoSyncKey(type)); } catch (_) {}
  }

  async function token() {
    const user = auth.currentUser;
    if (!user) throw new Error("Sessione scaduta.");
    return user.getIdToken();
  }

  async function call(payload) {
    const response = await fetch("/.netlify/functions/google-sheet-sync", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: `Bearer ${await token()}` },
      body: JSON.stringify(payload)
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok || !result.ok) throw new Error(result.error || "Sincronizzazione non riuscita.");
    return result;
  }

  function technical(record, source = "APP") {
    return { RECORD_ID: record.id, UPDATED_AT: timestamp(record.updatedAt || record.createdAt), UPDATED_BY: clean(record.updatedBy), SYNC_VERSION: Number(record.syncVersion) || 1, SYNC_SOURCE: source, ROW_STATUS: "ACTIVE" };
  }

  function makeSheets() {
    const api = window.HeraManagementV2;
    if (!api?.workbookRows) throw new Error("Modulo Personale e Mezzi non ancora disponibile. Riapri la sezione e riprova.");
    const personnel = api.workbookRows("personale", personaleRecords).map((row, index) => ({ ...row, ...technical(personaleRecords[index]) }));
    const vehicles = api.workbookRows("mezzi", mezziRecords).map((row, index) => ({ ...row, ...technical(mezziRecords[index]) }));
    const personnelLinks = personaleRecords.flatMap((person) => (HeraManagementCore.legacyEnabledIds(person, (value) => value) || []).map((commessaId) => ({ RECORD_ID: `${person.id}:${commessaId}`, ID_OPERATORE: person.id, ID_COMMESSA: commessaId, ROW_STATUS: "ACTIVE", UPDATED_AT: timestamp(person.updatedAt), UPDATED_BY: clean(person.updatedBy), SYNC_VERSION: Number(person.syncVersion) || 1, SYNC_SOURCE: "APP" })));
    const vehicleLinks = mezziRecords.flatMap((vehicle) => (Array.isArray(vehicle.commessaIds) ? vehicle.commessaIds : []).map((commessaId) => ({ RECORD_ID: `${vehicle.id}:${commessaId}`, ID_MEZZO: vehicle.id, ID_COMMESSA: commessaId, ROW_STATUS: "ACTIVE", UPDATED_AT: timestamp(vehicle.updatedAt), UPDATED_BY: clean(vehicle.updatedBy), SYNC_VERSION: Number(vehicle.syncVersion) || 1, SYNC_SOURCE: "APP" })));
    return { PERSONALE: personnel, MEZZI: vehicles, COMMESSE_PERSONALE: personnelLinks, COMMESSE_MEZZI: vehicleLinks, LOG_SINCRONIZZAZIONE: [] };
  }

  function autoSyncAllowed(type) {
    if (autoSyncChecked.has(type)) return false;
    autoSyncChecked.add(type);
    try {
      const last = Number(sessionStorage.getItem(autoSyncKey(type)) || 0);
      if (last && Date.now() - last < AUTO_SYNC_MIN_INTERVAL_MS) return false;
      sessionStorage.setItem(autoSyncKey(type), String(Date.now()));
    } catch (_) {}
    return true;
  }

  async function link(type) {
    const existing = await load(type);
    const supplied = window.prompt("Incolla il link di un Google Sheet esistente oppure lascia vuoto per crearne uno nuovo.", existing.sheetUrl || "");
    if (supplied === null) return;
    const result = await call({ action: "createRegistrySpreadsheet", registry: type, sheetUrl: clean(supplied), sheetNames: SHEETS });
    await save(type, {
      sheetUrl: result.sheetUrl,
      spreadsheetId: result.spreadsheetId,
      gid: result.gid || "0",
      legacyMode: result.legacyMode === true,
      linkedAt: firebase.firestore.FieldValue.serverTimestamp(),
      linkedBy: currentUser?.uid || "",
      autoSync: false,
      conflictPolicy: "LATEST_WINS",
      noAutomaticDeletion: true
    });
    alert(result.legacyMode
      ? "Google Sheet creato e collegato. La modalità compatibile è attiva e consente apertura e sincronizzazione dall’app verso il foglio."
      : "Google Sheet collegato. Puoi avviare la prima sincronizzazione.");
  }

  function sync(type) {
    if (syncInFlight.has(type)) return syncInFlight.get(type);
    const task = (async () => {
      const config = await load(type);
      if (!config.sheetUrl) return link(type);
      const result = await call({
        action: "syncRegistrySpreadsheet",
        registry: type,
        sheetUrl: config.sheetUrl,
        spreadsheetId: config.spreadsheetId,
        gid: config.gid || "0",
        sheets: makeSheets(),
        conflictPolicy: config.conflictPolicy || "LATEST_WINS",
        noAutomaticDeletion: true
      });
      await save(type, {
        lastSyncAt: firebase.firestore.FieldValue.serverTimestamp(),
        lastSyncBy: currentUser?.uid || "",
        lastConflictCount: result.conflicts?.length || 0,
        legacyMode: result.legacyMode === true
      });
      const primary = type === "personale" ? "PERSONALE" : "MEZZI";
      const idColumn = type === "personale" ? "ID_OPERATORE" : "ID_MEZZO";
      const incoming = (result.incoming?.[primary] || []).filter((row) => clean(row.ROW_STATUS).toUpperCase() !== "DELETED").map((row) => ({ ...row, [idColumn]: row[idColumn] || row.RECORD_ID }));
      if (incoming.length) {
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(incoming), primary);
        const bytes = XLSX.write(wb, { bookType: "xlsx", type: "array" });
        await window.HeraManagementV2.previewImport(type, { name: `Google_Sheets_${primary}.xlsx`, arrayBuffer: async () => bytes });
      } else if (result.legacyMode) {
        alert(`Sincronizzazione completata: ${result.rowsWritten || 0} righe inviate al Foglio Google.`);
      } else {
        alert(`Sincronizzazione completata. ${result.rowsWritten || 0} righe aggiornate; ${result.conflicts?.length || 0} conflitti; nessuna eliminazione automatica.`);
      }
    })().finally(() => syncInFlight.delete(type));
    syncInFlight.set(type, task);
    return task;
  }

  async function settings(type) {
    const config = await load(type);
    if (!config.sheetUrl) return link(type);
    const policy = window.prompt("Gestione conflitti: LATEST_WINS oppure APP_WINS.", config.conflictPolicy || "LATEST_WINS");
    if (policy === null) return;
    const normalized = clean(policy).toUpperCase();
    if (!["LATEST_WINS", "APP_WINS"].includes(normalized)) return alert("Impostazione non valida.");
    const autoSync = window.confirm("Abilitare la sincronizzazione automatica incrementale all’apertura della sezione?");
    await save(type, { conflictPolicy: normalized, autoSync, noAutomaticDeletion: true, updatedAt: firebase.firestore.FieldValue.serverTimestamp() });
    autoSyncChecked.delete(type);
    try { sessionStorage.removeItem(autoSyncKey(type)); } catch (_) {}
    alert("Impostazioni sincronizzazione salvate.");
  }

  async function unlink(type) {
    const config = await load(type);
    if (!config.sheetUrl) return;
    if (!window.confirm("Scollegare il Google Sheet? Il foglio e i dati non saranno eliminati.")) return;
    await remove(type);
    alert("Google Sheet scollegato senza eliminare dati.");
  }

  async function open(type) {
    const pendingTab = window.open("", "_blank");
    if (!pendingTab) throw new Error("Il browser ha bloccato l’apertura del Foglio Google. Consenti i popup per questa app.");
    try {
      pendingTab.opener = null;
      const config = await load(type);
      if (!config.sheetUrl) {
        pendingTab.close();
        return link(type);
      }
      const url = clean(config.sheetUrl);
      if (!/^https:\/\/docs\.google\.com\/spreadsheets\//i.test(url)) throw new Error("Il collegamento del Foglio Google non è valido. Collegalo nuovamente.");
      pendingTab.location.replace(url);
    } catch (error) {
      try { pendingTab.close(); } catch (_) {}
      throw error;
    }
  }

  const run = (operation) => void operation().catch((error) => alert(friendlyError(error)));

  function bind(root, type) {
    root.querySelector("[data-sheet-link]")?.addEventListener("click", () => run(() => link(type)));
    root.querySelector("[data-sheet-sync]")?.addEventListener("click", () => run(() => sync(type)));
    root.querySelector("[data-sheet-open]")?.addEventListener("click", () => run(() => open(type)));
    root.querySelector("[data-sheet-settings]")?.addEventListener("click", () => run(() => settings(type)));
    root.querySelector("[data-sheet-unlink]")?.addEventListener("click", () => run(() => unlink(type)));
    void load(type).then((config) => {
      if (config.autoSync && config.sheetUrl && navigator.onLine && autoSyncAllowed(type)) sync(type).catch(() => {});
    }).catch(() => {});
  }

  window.RegistryGoogleSheetSync = { bind, sync, SHEETS, TECHNICAL_COLUMNS };
})();

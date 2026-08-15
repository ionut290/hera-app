(() => {
  "use strict";

  const originalSelectCommessa = window.selectCommessa;
  if (typeof originalSelectCommessa !== "function" || originalSelectCommessa.__navigationRepair) return;

  function syncCommessaHeader(nome, codice = "") {
    const safeName = String(nome || "Commessa").trim() || "Commessa";
    const safeCode = String(codice || "").trim();
    const focusLabel = document.getElementById("commessa-focus-label");
    const focusCode = document.getElementById("commessa-focus-code");
    const pageTitle = document.getElementById("impianti-page-title");
    const activeLabel = document.getElementById("commessa-attiva");

    if (focusLabel) focusLabel.textContent = safeName.toUpperCase();
    if (focusCode) focusCode.textContent = safeCode;
    if (pageTitle) pageTitle.textContent = `Impianti commessa: ${safeName}`;
    if (activeLabel) {
      activeLabel.textContent = safeCode
        ? `Commessa selezionata: ${safeName} • Cod. commessa: ${safeCode}`
        : `Commessa selezionata: ${safeName}`;
    }
  }

  function forceCommessaNavigation(id, nome, codice = "") {
    syncCommessaHeader(nome, codice);
    try {
      localStorage.setItem("heraLastSelectedCommessaId", String(id || ""));
    } catch (_) {}

    try {
      window.stopImpiantiSubscription?.();
      window.stopCommessaNotesSubscription?.();

      const hasSubcommesse = typeof window.getSubcommesse === "function"
        && window.getSubcommesse(id).length > 0;

      if (!hasSubcommesse) {
        window.subscribeImpianti?.();
        window.subscribeCommessaNotes?.();
      }

      if (typeof window.setCommessaHash === "function") {
        window.setCommessaHash();
      } else {
        window.location.hash = `commessa=${encodeURIComponent(String(id || ""))}`;
      }
      window.applyRoute?.();
    } catch (fallbackError) {
      console.error("Ripristino apertura commessa non completato:", fallbackError);
      window.location.hash = `commessa=${encodeURIComponent(String(id || ""))}`;
    }
  }

  function selectCommessaWithNavigationRepair(id, nome, codice = "") {
    try {
      return originalSelectCommessa.call(this, id, nome, codice);
    } catch (error) {
      console.error("Errore durante apertura commessa; applico navigazione protetta:", {
        commessaId: id,
        commessaNome: nome,
        error
      });
      forceCommessaNavigation(id, nome, codice);
      return undefined;
    }
  }

  selectCommessaWithNavigationRepair.__navigationRepair = true;
  window.selectCommessa = selectCommessaWithNavigationRepair;
})();

// Cache persistente degli impianti: viene usata solo quando esiste già un
// checkpoint incrementale sicuro. Non sostituisce mai il fallback completo
// Firestore e non aggiunge letture, scritture o listener remoti.
(() => {
  "use strict";

  if (window.HeraImpiantiPersistentCache?.installed) return;
  if (
    typeof subscribeImpianti !== "function"
    || typeof renderImpiantiAfterRemoteSync !== "function"
    || typeof readImpiantiIncrementalState !== "function"
    || typeof saveImpiantiIncrementalState !== "function"
    || typeof impiantiByCommessaId === "undefined"
  ) return;

  const CACHE_SCHEMA_VERSION = 1;
  const CACHE_PREFIX = "heraImpiantiPersistentCacheV1:";
  const CACHE_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
  const CACHE_MAX_ENTRY_BYTES = 2 * 1024 * 1024;
  const CACHE_MAX_ENTRIES_PER_USER = 8;
  const contexts = new Map();
  const state = {
    hydrated: 0,
    persisted: 0,
    rejected: 0,
    removedExpired: 0,
    storageErrors: 0,
    lastCommessaId: ""
  };

  const originalSubscribeImpianti = subscribeImpianti;
  const originalRenderImpiantiAfterRemoteSync = renderImpiantiAfterRemoteSync;

  function getUserScope() {
    const uid = String(currentUser?.uid || "").trim();
    if (!uid) return null;
    let collectionName = "commesse";
    try {
      if (typeof getCommesseCollectionName === "function") {
        collectionName = String(getCommesseCollectionName() || "commesse").trim() || "commesse";
      }
    } catch (_) {}
    return { uid, collectionName };
  }

  function getCacheKey(commessaId, scope = getUserScope()) {
    if (!scope || !commessaId) return "";
    return `${CACHE_PREFIX}${encodeURIComponent(scope.uid)}:${encodeURIComponent(scope.collectionName)}:${encodeURIComponent(commessaId)}`;
  }

  function cloneItems(items) {
    return Array.isArray(items) ? items.map((item) => ({ ...(item || {}) })) : [];
  }

  function removeCacheKey(key, expired = false) {
    if (!key) return;
    try {
      localStorage.removeItem(key);
      if (expired) state.removedExpired += 1;
    } catch (_) {
      state.storageErrors += 1;
    }
  }

  function readPersistentCache(commessaId) {
    const scope = getUserScope();
    const key = getCacheKey(commessaId, scope);
    if (!scope || !key) return null;
    try {
      const parsed = JSON.parse(localStorage.getItem(key) || "null");
      const ageMs = Date.now() - Number(parsed?.savedAt || 0);
      const valid = Boolean(
        parsed
        && parsed.schemaVersion === CACHE_SCHEMA_VERSION
        && parsed.uid === scope.uid
        && parsed.collectionName === scope.collectionName
        && parsed.commessaId === commessaId
        && Number(parsed.markerMs) > 0
        && Array.isArray(parsed.items)
        && parsed.items.length > 0
        && Number.isFinite(ageMs)
        && ageMs >= 0
        && ageMs <= CACHE_MAX_AGE_MS
      );
      if (!valid) {
        if (parsed) removeCacheKey(key, ageMs > CACHE_MAX_AGE_MS);
        state.rejected += 1;
        return null;
      }
      return { ...parsed, key };
    } catch (_) {
      state.storageErrors += 1;
      removeCacheKey(key);
      return null;
    }
  }

  function prunePersistentCaches(scope, keepKey) {
    if (!scope) return;
    try {
      const scopedPrefix = `${CACHE_PREFIX}${encodeURIComponent(scope.uid)}:${encodeURIComponent(scope.collectionName)}:`;
      const entries = [];
      for (let index = 0; index < localStorage.length; index += 1) {
        const key = localStorage.key(index);
        if (!key || !key.startsWith(scopedPrefix)) continue;
        let savedAt = 0;
        try {
          savedAt = Number(JSON.parse(localStorage.getItem(key) || "null")?.savedAt || 0);
        } catch (_) {}
        entries.push({ key, savedAt });
      }
      entries.sort((a, b) => b.savedAt - a.savedAt);
      entries.slice(CACHE_MAX_ENTRIES_PER_USER).forEach(({ key }) => {
        if (key !== keepKey) removeCacheKey(key);
      });
    } catch (_) {
      state.storageErrors += 1;
    }
  }

  function persistCurrentCache(commessaId) {
    const scope = getUserScope();
    const markerMs = Number(readImpiantiIncrementalState(commessaId)?.lastChangedAtMs || 0);
    const items = cloneItems(currentImpianti);
    const key = getCacheKey(commessaId, scope);
    if (!scope || !key || markerMs <= 0 || !items.length) return false;

    const payload = {
      schemaVersion: CACHE_SCHEMA_VERSION,
      uid: scope.uid,
      collectionName: scope.collectionName,
      commessaId,
      markerMs,
      savedAt: Date.now(),
      items
    };

    try {
      const serialized = JSON.stringify(payload);
      if (serialized.length > CACHE_MAX_ENTRY_BYTES) {
        state.rejected += 1;
        removeCacheKey(key);
        return false;
      }
      localStorage.setItem(key, serialized);
      state.persisted += 1;
      state.lastCommessaId = commessaId;
      prunePersistentCaches(scope, key);
      return true;
    } catch (_) {
      state.storageErrors += 1;
      return false;
    }
  }

  function hydratePersistentCache(commessaId) {
    const cached = readPersistentCache(commessaId);
    if (!cached) return false;
    try {
      impiantiByCommessaId.set(commessaId, cloneItems(cached.items));
      saveImpiantiIncrementalState(commessaId, Number(cached.markerMs));
      state.hydrated += 1;
      state.lastCommessaId = commessaId;
      return true;
    } catch (_) {
      state.rejected += 1;
      removeCacheKey(cached.key);
      return false;
    }
  }

  subscribeImpianti = function subscribeImpiantiWithPersistentCache() {
    const commessaId = String(selectedCommessaId || "").trim();
    if (!commessaId) return originalSubscribeImpianti.apply(this, arguments);

    const existing = impiantiByCommessaId.get(commessaId);
    const hadMemoryCache = Array.isArray(existing) && existing.length > 0;
    const hydrated = hadMemoryCache ? false : hydratePersistentCache(commessaId);
    const checkpoint = readImpiantiIncrementalState(commessaId);

    contexts.set(commessaId, {
      canPersist: Boolean((hadMemoryCache || hydrated) && Number(checkpoint?.lastChangedAtMs) > 0),
      hydrated,
      startedAt: Date.now()
    });

    return originalSubscribeImpianti.apply(this, arguments);
  };

  renderImpiantiAfterRemoteSync = function renderImpiantiAfterRemoteSyncWithPersistentCache(rawImpianti, previousDoneSignatureRef) {
    const commessaId = String(selectedCommessaId || "").trim();
    const result = originalRenderImpiantiAfterRemoteSync.apply(this, arguments);
    const context = contexts.get(commessaId);

    // Si salva solo dopo un ciclo partito da cache + checkpoint. Il primo
    // caricamento completo non viene mai promosso direttamente a cache persistente:
    // evita finestre di gara tra snapshot completo e marker del server.
    if (commessaId && context?.canPersist) {
      persistCurrentCache(commessaId);
    }
    return result;
  };

  window.HeraImpiantiPersistentCache = {
    installed: true,
    version: "1.0.0",
    mode: "verified-incremental-only",
    maxAgeMs: CACHE_MAX_AGE_MS,
    getState: () => ({ ...state, activeContexts: contexts.size }),
    clearCurrent: () => {
      const commessaId = String(selectedCommessaId || "").trim();
      removeCacheKey(getCacheKey(commessaId));
    }
  };
})();

// Le statistiche degli impianti di tutte le commesse non devono aprire un
// listener completo all'avvio/Home. La diagnostica ha mostrato che questo
// percorso rilegge intere sottocollezioni anche quando l'utente non sta
// consultando un riepilogo. Il listener originale resta disponibile e viene
// attivato solo quando si apre una commessa padre, dove i rollup delle
// subcommesse sono effettivamente necessari.
(() => {
  "use strict";

  if (window.HeraLazyCommessaStats?.installed) return;
  if (
    typeof refreshCommesseDependentUI !== "function"
    || typeof subscribeStatsForCommesse !== "function"
    || typeof selectCommessa !== "function"
  ) return;

  const originalRefreshCommesseDependentUI = refreshCommesseDependentUI;
  const originalSubscribeStatsForCommesse = subscribeStatsForCommesse;
  const originalSelectCommessaForStats = selectCommessa;
  const state = {
    suppressedStartupLoads: 0,
    explicitParentLoads: 0,
    lastParentId: "",
    errors: []
  };

  refreshCommesseDependentUI = function refreshCommesseDependentUILazyStats(includeRemoteStats = true) {
    if (includeRemoteStats) state.suppressedStartupLoads += 1;
    // Mantiene rendering e cache commesse invariati, ma non apre automaticamente
    // i listener impianti di tutte le commesse né il listener ore completo.
    return originalRefreshCommesseDependentUI.call(this, false);
  };

  selectCommessa = function selectCommessaWithLazyStats(id, nome, codice = "") {
    const result = originalSelectCommessaForStats.call(this, id, nome, codice);
    try {
      const hasSubcommesse = typeof getSubcommesse === "function" && getSubcommesse(id).length > 0;
      if (hasSubcommesse) {
        state.explicitParentLoads += 1;
        state.lastParentId = String(id || "");
        originalSubscribeStatsForCommesse();
      }
    } catch (error) {
      state.errors.push(String(error?.message || error));
      console.warn("Statistiche commessa padre non caricate:", error);
    }
    return result;
  };

  window.HeraLazyCommessaStats = {
    installed: true,
    version: "1.0.0",
    mode: "home-suppressed-parent-explicit",
    loadAll: () => originalSubscribeStatsForCommesse(),
    getState: () => ({ ...state, errors: state.errors.slice() })
  };
})();

// Evita il ciclo stop -> subscribe immediato sulla stessa commessa, che causa
// una seconda consegna completa degli stessi impianti. La chiusura viene
// differita solo se esiste già un listener impianti attivo; se cambia commessa
// o non arriva una riapertura immediata, la chiusura originale viene eseguita.
// Non modifica FATTO, fattoVisualEvidence, note commessa o WhatsApp/WHAZZUP.
(() => {
  "use strict";

  if (window.HeraImpiantiListenerLifecycleGuard?.installed) return;
  if (
    typeof window.stopImpiantiSubscription !== "function"
    || typeof window.subscribeImpianti !== "function"
  ) return;

  const originalStopImpiantiSubscription = window.stopImpiantiSubscription;
  const originalSubscribeImpianti = window.subscribeImpianti;
  const DEFER_MS = 80;
  let pendingStop = null;
  let activeCommessaId = "";

  const state = {
    deferredStops: 0,
    cancelledSameCommessaRestarts: 0,
    flushedForCommessaChange: 0,
    completedStops: 0,
    subscribes: 0
  };

  function selectedId() {
    try {
      return String(window.selectedCommessaId || selectedCommessaId || "").trim();
    } catch (_) {
      return String(window.selectedCommessaId || "").trim();
    }
  }

  function clearPendingTimer() {
    if (pendingStop?.timer) clearTimeout(pendingStop.timer);
  }

  function runPendingStop() {
    if (!pendingStop) return;
    const stop = pendingStop;
    pendingStop = null;
    clearTimeout(stop.timer);
    originalStopImpiantiSubscription.call(window);
    if (activeCommessaId === stop.commessaId) activeCommessaId = "";
    state.completedStops += 1;
  }

  window.stopImpiantiSubscription = function stopImpiantiSubscriptionWithLifecycleGuard() {
    if (!activeCommessaId) {
      return originalStopImpiantiSubscription.apply(this, arguments);
    }

    if (pendingStop) {
      if (pendingStop.commessaId === activeCommessaId) return;
      runPendingStop();
    }

    const commessaId = activeCommessaId;
    const timer = setTimeout(() => {
      if (pendingStop?.commessaId !== commessaId) return;
      runPendingStop();
    }, DEFER_MS);
    pendingStop = { commessaId, timer, createdAt: Date.now() };
    state.deferredStops += 1;
    return undefined;
  };

  window.subscribeImpianti = function subscribeImpiantiWithLifecycleGuard() {
    const commessaId = selectedId();
    if (pendingStop) {
      const sameCommessa = Boolean(
        commessaId
        && activeCommessaId
        && pendingStop.commessaId === commessaId
        && activeCommessaId === commessaId
      );
      if (sameCommessa) {
        clearPendingTimer();
        pendingStop = null;
        state.cancelledSameCommessaRestarts += 1;
        return undefined;
      }
      state.flushedForCommessaChange += 1;
      runPendingStop();
    }

    const result = originalSubscribeImpianti.apply(this, arguments);
    activeCommessaId = commessaId;
    state.subscribes += 1;
    return result;
  };

  window.HeraImpiantiListenerLifecycleGuard = {
    installed: true,
    version: "1.0.0",
    deferMs: DEFER_MS,
    getState: () => ({
      ...state,
      activeCommessaId,
      pendingStopCommessaId: pendingStop?.commessaId || ""
    }),
    flush: runPendingStop
  };
})();

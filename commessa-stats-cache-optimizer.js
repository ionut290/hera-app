(() => {
  "use strict";

  if (window.HeraCommessaStatsCacheOptimizer?.installed) return;
  if (
    typeof db === "undefined" || !db
    || typeof calculateImpiantiStats !== "function"
    || typeof combineImpiantiForView !== "function"
    || typeof recalculateCommessaWorkSummaries !== "function"
    || typeof subscribeStatsForCommesse !== "function"
    || typeof commesseById === "undefined"
    || typeof impiantiByCommessaId === "undefined"
    || typeof commessaStatsById === "undefined"
    || typeof unsubscribeCommessaStats === "undefined"
  ) return;

  const CACHE_VERSION = 1;
  const CACHE_PREFIX = "heraCommessaStatsCacheV1:";
  const IMPIANTI_CACHE_PREFIX = "heraImpiantiPersistentCacheV1:";
  const MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
  const originalSubscribeStatsForCommesse = subscribeStatsForCommesse;
  const state = {
    cacheHits: 0,
    cacheMisses: 0,
    fullFallbackLoads: 0,
    incrementalListeners: 0,
    incrementalDeliveries: 0,
    changedDocumentsRead: 0,
    errors: []
  };

  function getScope() {
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

  function activeCollectionName() {
    return getScope()?.collectionName || "commesse";
  }

  function commessaRef(commessaId) {
    return db.collection(activeCollectionName()).doc(commessaId);
  }

  function cacheKey(commessaId, scope = getScope()) {
    if (!scope || !commessaId) return "";
    return `${CACHE_PREFIX}${encodeURIComponent(scope.uid)}:${encodeURIComponent(scope.collectionName)}:${encodeURIComponent(commessaId)}`;
  }

  function impiantiCacheKey(commessaId, scope = getScope()) {
    if (!scope || !commessaId) return "";
    return `${IMPIANTI_CACHE_PREFIX}${encodeURIComponent(scope.uid)}:${encodeURIComponent(scope.collectionName)}:${encodeURIComponent(commessaId)}`;
  }

  function cloneItems(items) {
    return Array.isArray(items) ? items.map((item) => ({ ...(item || {}) })) : [];
  }

  function timestampMs(value) {
    if (!value) return 0;
    if (typeof value.toMillis === "function") return Number(value.toMillis()) || 0;
    if (typeof value.toDate === "function") return Number(value.toDate()?.getTime?.()) || 0;
    if (value instanceof Date) return value.getTime();
    return Number(value) || 0;
  }

  function readJson(key) {
    if (!key) return null;
    try {
      return JSON.parse(localStorage.getItem(key) || "null");
    } catch (_) {
      return null;
    }
  }

  function readImpiantiCache(commessaId) {
    const scope = getScope();
    const parsed = readJson(impiantiCacheKey(commessaId, scope));
    const age = Date.now() - Number(parsed?.savedAt || 0);
    if (!scope || !parsed || parsed.schemaVersion !== 1 || parsed.uid !== scope.uid || parsed.collectionName !== scope.collectionName
      || parsed.commessaId !== commessaId || !Array.isArray(parsed.items) || !parsed.items.length
      || !Number(parsed.markerMs) || age < 0 || age > MAX_AGE_MS) return null;
    return { items: cloneItems(parsed.items), markerMs: Number(parsed.markerMs) };
  }

  function readStatsCache(commessaId) {
    const scope = getScope();
    const parsed = readJson(cacheKey(commessaId, scope));
    const age = Date.now() - Number(parsed?.savedAt || 0);
    if (!scope || !parsed || parsed.version !== CACHE_VERSION || parsed.uid !== scope.uid || parsed.collectionName !== scope.collectionName
      || parsed.commessaId !== commessaId || !parsed.stats || !Number(parsed.markerMs)
      || age < 0 || age > MAX_AGE_MS) return null;
    return parsed;
  }

  function writeStatsCache(commessaId, items, markerMs) {
    const scope = getScope();
    const key = cacheKey(commessaId, scope);
    if (!scope || !key || !Array.isArray(items) || !items.length || !Number(markerMs)) return false;
    const rawItems = cloneItems(items);
    const stats = calculateImpiantiStats(rawItems);
    const payload = {
      version: CACHE_VERSION,
      uid: scope.uid,
      collectionName: scope.collectionName,
      commessaId,
      markerMs: Number(markerMs),
      savedAt: Date.now(),
      stats
    };
    try {
      localStorage.setItem(key, JSON.stringify(payload));
      return true;
    } catch (_) {
      return false;
    }
  }

  function refreshStatsUI() {
    recalculateCommessaWorkSummaries();
    try { renderCommesseHomeList?.(); } catch (_) {}
    try { renderCommesseManagementList?.(); } catch (_) {}
    try { renderParentCommessaOverview?.(); } catch (_) {}
    try { updateCommessaDashboard?.(); } catch (_) {}
  }

  function applyItems(commessaId, items, markerMs = 0) {
    const rawItems = cloneItems(items);
    const combined = combineImpiantiForView(rawItems);
    impiantiByCommessaId.set(commessaId, combined);
    commessaStatsById.set(commessaId, calculateImpiantiStats(rawItems));
    if (markerMs > 0) writeStatsCache(commessaId, rawItems, markerMs);
    refreshStatsUI();
  }

  function targetCommessaIds() {
    const selected = String(selectedCommessaId || "").trim();
    if (!selected || typeof getSubcommesse !== "function") return [];
    const children = getSubcommesse(selected) || [];
    return children.map((item) => String(item?.id || "").trim()).filter(Boolean);
  }

  function stopUnused(targetIds) {
    const targetSet = new Set(targetIds);
    Array.from(unsubscribeCommessaStats.keys()).forEach((commessaId) => {
      if (!targetSet.has(commessaId)) {
        try { unsubscribeCommessaStats.get(commessaId)?.(); } catch (_) {}
        unsubscribeCommessaStats.delete(commessaId);
      }
    });
  }

  async function latestMarker(commessaId) {
    try {
      const snapshot = await commessaRef(commessaId).collection("impiantoChangeIndex")
        .orderBy("changedAt", "desc").limit(1).get();
      if (snapshot.empty) return 0;
      return timestampMs(snapshot.docs[0]?.data?.()?.changedAt);
    } catch (error) {
      state.errors.push(`marker ${commessaId}: ${String(error?.message || error)}`);
      return 0;
    }
  }

  async function readChangedDocument(commessaId, impiantoId) {
    const ref = commessaRef(commessaId).collection("impianti").doc(impiantoId);
    const snap = await ref.get();
    state.changedDocumentsRead += snap.exists ? 1 : 0;
    return snap.exists ? { id: snap.id, ...snap.data() } : null;
  }

  function startIncrementalListener(commessaId, markerMs) {
    if (!markerMs || unsubscribeCommessaStats.has(commessaId)) return;
    const markerDate = new Date(markerMs);
    const query = commessaRef(commessaId).collection("impiantoChangeIndex")
      .where("changedAt", ">", markerDate)
      .orderBy("changedAt", "asc");

    const unsubscribe = query.onSnapshot(async (snapshot) => {
      if (snapshot.empty) return;
      const deltaDocs = typeof snapshot.docChanges === "function"
        ? snapshot.docChanges().filter((change) => change.type !== "removed").map((change) => change.doc)
        : snapshot.docs;
      if (!deltaDocs.length) return;
      state.incrementalDeliveries += 1;
      try {
        const changes = deltaDocs.map((doc) => ({ id: doc.id, markerMs: timestampMs(doc.data()?.changedAt) }));
        const current = cloneItems(impiantiByCommessaId.get(commessaId) || []);
        const byId = new Map(current.map((item) => [String(item.id || ""), item]));
        let newestMarker = markerMs;
        for (const change of changes) {
          if (!change.id) continue;
          const changed = await readChangedDocument(commessaId, change.id);
          if (changed) byId.set(change.id, changed);
          else byId.delete(change.id);
          newestMarker = Math.max(newestMarker, change.markerMs || 0);
        }
        const items = Array.from(byId.values());
        applyItems(commessaId, items, newestMarker);
        startIncrementalListener.lastMarkers.set(commessaId, newestMarker);
      } catch (error) {
        state.errors.push(`incremental ${commessaId}: ${String(error?.message || error)}`);
        console.warn("Aggiornamento incrementale statistiche non riuscito:", error);
      }
    }, (error) => {
      state.errors.push(`listener ${commessaId}: ${String(error?.message || error)}`);
      console.warn("Listener incrementale statistiche non disponibile:", error);
    });
    unsubscribeCommessaStats.set(commessaId, unsubscribe);
    state.incrementalListeners += 1;
  }
  startIncrementalListener.lastMarkers = new Map();

  async function bootstrapCommessa(commessaId) {
    if (!commessaId || unsubscribeCommessaStats.has(commessaId)) return;

    const memoryItems = cloneItems(impiantiByCommessaId.get(commessaId) || []);
    const persistentItems = memoryItems.length ? null : readImpiantiCache(commessaId);
    const statsCache = readStatsCache(commessaId);

    if (memoryItems.length) {
      const marker = Number(statsCache?.markerMs || persistentItems?.markerMs || 0);
      commessaStatsById.set(commessaId, calculateImpiantiStats(memoryItems));
      if (marker) writeStatsCache(commessaId, memoryItems, marker);
      state.cacheHits += 1;
      refreshStatsUI();
      if (marker) startIncrementalListener(commessaId, marker);
      else await fallbackFullLoad(commessaId);
      return;
    }

    if (persistentItems?.items?.length && persistentItems.markerMs) {
      state.cacheHits += 1;
      applyItems(commessaId, persistentItems.items, persistentItems.markerMs);
      startIncrementalListener(commessaId, persistentItems.markerMs);
      return;
    }

    if (statsCache?.stats) {
      // La sola cache statistiche permette il rendering immediato, ma per applicare
      // modifiche incrementali serve la lista impianti. Non viene mai trattata come
      // fonte sufficiente per evitare il fallback completo.
      commessaStatsById.set(commessaId, statsCache.stats);
      refreshStatsUI();
    }

    state.cacheMisses += 1;
    await fallbackFullLoad(commessaId);
  }

  async function fallbackFullLoad(commessaId) {
    state.fullFallbackLoads += 1;
    const markerBefore = await latestMarker(commessaId);
    try {
      const snapshot = await commessaRef(commessaId).collection("impianti").get();
      const rawItems = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      applyItems(commessaId, rawItems, markerBefore);
      if (markerBefore > 0) startIncrementalListener(commessaId, markerBefore);
      else {
        // Indice non ancora disponibile: manteniamo il comportamento storico come
        // fallback di sicurezza, senza trasformare una cache non verificabile in fonte primaria.
        const fallbackUnsubscribe = commessaRef(commessaId).collection("impianti").onSnapshot((live) => {
          const liveItems = live.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
          applyItems(commessaId, liveItems, 0);
        });
        unsubscribeCommessaStats.set(commessaId, fallbackUnsubscribe);
      }
    } catch (error) {
      state.errors.push(`fallback ${commessaId}: ${String(error?.message || error)}`);
      console.warn("Caricamento statistiche commessa non riuscito:", error);
    }
  }

  subscribeStatsForCommesse = function subscribeStatsForCommesseCachedIncremental() {
    const targets = targetCommessaIds();
    stopUnused(targets);
    if (!targets.length) return;
    targets.forEach((commessaId) => {
      bootstrapCommessa(commessaId).catch((error) => {
        state.errors.push(`bootstrap ${commessaId}: ${String(error?.message || error)}`);
      });
    });
  };

  window.HeraCommessaStatsCacheOptimizer = {
    installed: true,
    version: "1.0.0",
    mode: "persistent-cache-plus-change-index",
    originalSubscribeStatsForCommesse,
    getState: () => ({ ...state, errors: state.errors.slice(), active: unsubscribeCommessaStats.size })
  };
})();

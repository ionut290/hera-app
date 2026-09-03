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
  const TAP_RELEASE_GRACE_MS = 180;
  const originalSubscribeStatsForCommesse = subscribeStatsForCommesse;

  let statsUiRefreshScheduled = false;
  let statsUiRefreshPendingWhileHidden = false;
  let interactionGuardUntil = 0;
  let interactionPointerDown = false;
  let deferredInteractionRefreshTimer = null;

  const state = {
    cacheHits: 0,
    cacheMisses: 0,
    fullFallbackLoads: 0,
    incrementalListeners: 0,
    incrementalDeliveries: 0,
    changedDocumentsRead: 0,
    deferredUiRefreshes: 0,
    deferredInteractionRefreshes: 0,
    tapCacheWarmups: 0,
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
    if (
      !scope || !parsed || parsed.schemaVersion !== 1
      || parsed.uid !== scope.uid || parsed.collectionName !== scope.collectionName
      || parsed.commessaId !== commessaId || !Array.isArray(parsed.items) || !parsed.items.length
      || !Number(parsed.markerMs) || age < 0 || age > MAX_AGE_MS
    ) return null;
    return { items: cloneItems(parsed.items), markerMs: Number(parsed.markerMs) };
  }

  function readStatsCache(commessaId) {
    const scope = getScope();
    const parsed = readJson(cacheKey(commessaId, scope));
    const age = Date.now() - Number(parsed?.savedAt || 0);
    if (
      !scope || !parsed || parsed.version !== CACHE_VERSION
      || parsed.uid !== scope.uid || parsed.collectionName !== scope.collectionName
      || parsed.commessaId !== commessaId || !parsed.stats || !Number(parsed.markerMs)
      || age < 0 || age > MAX_AGE_MS
    ) return null;
    return parsed;
  }

  function writeStatsCache(commessaId, items, markerMs) {
    const scope = getScope();
    const key = cacheKey(commessaId, scope);
    if (!scope || !key || !Array.isArray(items) || !items.length || !Number(markerMs)) return false;
    const rawItems = cloneItems(items);
    try {
      localStorage.setItem(key, JSON.stringify({
        version: CACHE_VERSION,
        uid: scope.uid,
        collectionName: scope.collectionName,
        commessaId,
        markerMs: Number(markerMs),
        savedAt: Date.now(),
        stats: calculateImpiantiStats(rawItems)
      }));
      return true;
    } catch (_) {
      return false;
    }
  }

  function pageIsHidden() {
    return typeof document !== "undefined" && document.visibilityState === "hidden";
  }

  function isProtectedTapTarget(target) {
    if (!(target instanceof Element)) return false;
    return Boolean(target.closest(
      "#commesse-lista, #today-squads-section, #today-summary-card, #commesse-manage-list"
    ));
  }

  function getTappedCommessaId(target) {
    if (!(target instanceof Element)) return "";
    const node = target.closest("[data-commessa-id], .commessa-btn");
    return String(node?.dataset?.commessaId || "").trim();
  }

  function warmTappedCommessaCache(target) {
    const commessaId = getTappedCommessaId(target);
    if (!commessaId) return false;
    const existing = cloneItems(impiantiByCommessaId.get(commessaId) || []);
    if (existing.length) return false;
    const persistent = readImpiantiCache(commessaId);
    if (!persistent?.items?.length) return false;

    const rawItems = cloneItems(persistent.items);
    impiantiByCommessaId.set(commessaId, combineImpiantiForView(rawItems));
    commessaStatsById.set(commessaId, calculateImpiantiStats(rawItems));
    state.tapCacheWarmups += 1;
    return true;
  }

  function interactionGuardActive() {
    return interactionPointerDown || Date.now() < interactionGuardUntil;
  }

  function scheduleRefreshAfterInteraction() {
    if (deferredInteractionRefreshTimer) return;
    state.deferredInteractionRefreshes += 1;
    const delay = interactionPointerDown
      ? TAP_RELEASE_GRACE_MS
      : Math.max(16, interactionGuardUntil - Date.now() + 16);
    deferredInteractionRefreshTimer = window.setTimeout(() => {
      deferredInteractionRefreshTimer = null;
      if (interactionGuardActive()) {
        scheduleRefreshAfterInteraction();
        return;
      }
      refreshStatsUI();
    }, delay);
  }

  function onProtectedPointerDown(event) {
    if (!isProtectedTapTarget(event?.target)) return;
    if (typeof event.button === "number" && event.button !== 0) return;
    warmTappedCommessaCache(event.target);
    interactionPointerDown = true;
    interactionGuardUntil = Number.POSITIVE_INFINITY;
  }

  function onProtectedPointerRelease(event) {
    if (!interactionPointerDown && !isProtectedTapTarget(event?.target)) return;
    interactionPointerDown = false;
    interactionGuardUntil = Date.now() + TAP_RELEASE_GRACE_MS;
  }

  function refreshStatsUI() {
    if (pageIsHidden()) {
      if (!statsUiRefreshPendingWhileHidden) state.deferredUiRefreshes += 1;
      statsUiRefreshPendingWhileHidden = true;
      return;
    }
    if (interactionGuardActive()) {
      scheduleRefreshAfterInteraction();
      return;
    }
    if (statsUiRefreshScheduled) return;
    statsUiRefreshScheduled = true;
    const run = () => {
      statsUiRefreshScheduled = false;
      if (pageIsHidden()) {
        if (!statsUiRefreshPendingWhileHidden) state.deferredUiRefreshes += 1;
        statsUiRefreshPendingWhileHidden = true;
        return;
      }
      if (interactionGuardActive()) {
        scheduleRefreshAfterInteraction();
        return;
      }
      statsUiRefreshPendingWhileHidden = false;
      recalculateCommessaWorkSummaries();
      try { renderCommesseHomeList?.(); } catch (_) {}
      try { renderCommesseManagementList?.(); } catch (_) {}
      try { renderParentCommessaOverview?.(); } catch (_) {}
      try { updateCommessaDashboard?.(); } catch (_) {}
    };
    if (typeof window.requestAnimationFrame === "function") window.requestAnimationFrame(run);
    else window.setTimeout(run, 0);
  }

  if (typeof document !== "undefined" && typeof document.addEventListener === "function") {
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible" && statsUiRefreshPendingWhileHidden) refreshStatsUI();
    });
    document.addEventListener("pointerdown", onProtectedPointerDown, { capture: true, passive: true });
    document.addEventListener("pointerup", onProtectedPointerRelease, { capture: true, passive: true });
    document.addEventListener("pointercancel", onProtectedPointerRelease, { capture: true, passive: true });
  }

  function applyItems(commessaId, items, markerMs = 0) {
    const rawItems = cloneItems(items);
    impiantiByCommessaId.set(commessaId, combineImpiantiForView(rawItems));
    commessaStatsById.set(commessaId, calculateImpiantiStats(rawItems));
    if (markerMs > 0) writeStatsCache(commessaId, rawItems, markerMs);
    refreshStatsUI();
  }

  function targetCommessaIds() {
    const selected = String(selectedCommessaId || "").trim();
    if (!selected) return [];
    const children = typeof getSubcommesse === "function" ? (getSubcommesse(selected) || []) : [];
    return Array.from(new Set([
      selected,
      ...children.map((item) => String(item?.id || "").trim()).filter(Boolean)
    ]));
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
    const snap = await commessaRef(commessaId).collection("impianti").doc(impiantoId).get();
    state.changedDocumentsRead += snap.exists ? 1 : 0;
    return snap.exists ? { id: snap.id, ...snap.data() } : null;
  }

  function startIncrementalListener(commessaId, markerMs) {
    if (!markerMs || unsubscribeCommessaStats.has(commessaId)) return;
    const query = commessaRef(commessaId).collection("impiantoChangeIndex")
      .where("changedAt", ">", new Date(markerMs))
      .orderBy("changedAt", "asc");

    const unsubscribe = query.onSnapshot(async (snapshot) => {
      if (snapshot.empty) return;
      const deltaDocs = typeof snapshot.docChanges === "function"
        ? snapshot.docChanges().filter((change) => change.type !== "removed").map((change) => change.doc)
        : snapshot.docs;
      if (!deltaDocs.length) return;
      state.incrementalDeliveries += 1;
      try {
        const current = cloneItems(impiantiByCommessaId.get(commessaId) || []);
        const byId = new Map(current.map((item) => [String(item.id || ""), item]));
        let newestMarker = markerMs;
        for (const doc of deltaDocs) {
          const id = String(doc.id || "").trim();
          if (!id) continue;
          const changed = await readChangedDocument(commessaId, id);
          if (changed) byId.set(id, changed);
          else byId.delete(id);
          newestMarker = Math.max(newestMarker, timestampMs(doc.data()?.changedAt) || 0);
        }
        applyItems(commessaId, Array.from(byId.values()), newestMarker);
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
      const marker = Number(statsCache?.markerMs || 0);
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
      if (markerBefore > 0) {
        startIncrementalListener(commessaId, markerBefore);
      } else {
        const fallbackUnsubscribe = commessaRef(commessaId).collection("impianti").onSnapshot((live) => {
          applyItems(commessaId, live.docs.map((doc) => ({ id: doc.id, ...doc.data() })), 0);
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
    version: "1.5.0",
    mode: "tap-warmup-selected-first-persistent-cache-plus-change-index",
    originalSubscribeStatsForCommesse,
    getState: () => ({
      ...state,
      errors: state.errors.slice(),
      active: unsubscribeCommessaStats.size,
      uiRefreshPendingWhileHidden: statsUiRefreshPendingWhileHidden,
      interactionGuardActive: interactionGuardActive()
    })
  };
})();
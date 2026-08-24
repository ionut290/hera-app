/* INRETE Modena: usa soltanto la commessa canonica Cod. 28015 e mantiene separati impianti fisici e lavorazioni. */
(() => {
  "use strict";

  const SYNTHETIC_COMMESSA_ID = "inrete_modena_agosto_2026";
  const CANONICAL_COMMESSA_CODE = "28015";
  const CANONICAL_COMMESSA_NAME = "INRETE MODENA";
  const SUMMARY_REFRESH_MS = 30000;
  const LOCAL_GUARD_MS = 1500;
  const MANAGEMENT_REFRESH_MS = 800;
  const MANAGEMENT_DATA_REFRESH_MS = 60000;

  let summaryPromise = null;
  let lastSummaryAt = 0;
  let lastSummary = null;
  let duplicateArchivePromise = null;
  const managementSummaryById = new Map();
  const managementSummaryAtById = new Map();
  const managementSummaryPromiseById = new Map();

  const text = (value) => String(value ?? "").trim();
  const norm = (value) => text(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("it-IT")
    .replace(/[^a-z0-9]+/g, "");
  const normalizeCode = (value) => text(value).toLocaleUpperCase("it-IT").replace(/[^A-Z0-9]+/g, "");
  const status = (value) => text(value).toLocaleUpperCase("it-IT").replace(/_/g, " ");
  const isDone = (item = {}) => item.done === true || ["FATTO", "DONE", "COMPLETATO", "COMPLETATA", "ESEGUITO", "ESEGUITA"].includes(status(item.statoGenerale || item.stato));
  const isTodo = (item = {}) => !isDone(item) && ["", "DA FARE", "TODO", "APERTO", "APERTA"].includes(status(item.statoGenerale || item.stato));

  function activeCollectionName() {
    try {
      if (typeof getCommesseCollectionName === "function") return text(getCommesseCollectionName()) || "commesse";
    } catch (_) {}
    return "commesse";
  }

  function canWriteParentCommessa() {
    try {
      return typeof canManageData === "function" && canManageData() === true;
    } catch (_) {
      return false;
    }
  }

  function findCanonicalModenaCommessa() {
    try {
      if (activeCollectionName() !== "commesse" || typeof commesseById === "undefined" || !commesseById?.entries) return null;
      let nameMatch = null;
      for (const [id, raw] of commesseById.entries()) {
        const commessa = { id, ...(raw || {}) };
        if (!id || id === SYNTHETIC_COMMESSA_ID || commessa.hiddenFromHome === true || commessa.mergedIntoCommessaId) continue;
        const code = normalizeCode(commessa.codice || commessa.code);
        const name = norm(commessa.nome || commessa.name);
        if (code === CANONICAL_COMMESSA_CODE) return commessa;
        if (name === norm(CANONICAL_COMMESSA_NAME)) nameMatch = nameMatch || commessa;
      }
      return nameMatch;
    } catch (_) {
      return null;
    }
  }

  function physicalIdentity(row = {}, plantIdToKey = null) {
    const sap = normalizeCode(row.idSap || row.idSAP || row["ID SAP"] || row.codiceSap);
    if (sap) return `SAP:${sap}`;

    const linkedId = text(row.impiantoId || row.physicalPlantId || row.migrationSourceId).split("::")[0];
    if (linkedId && plantIdToKey?.has(linkedId)) return plantIdToKey.get(linkedId);
    if (linkedId) return `ID:${linkedId}`;

    const docId = text(row.id);
    if (docId && plantIdToKey?.has(docId)) return plantIdToKey.get(docId);
    if (docId) return `ID:${docId}`;

    const anag = norm(`${row.denominazione || row.nome || ""}|${row.comune || ""}|${row.indirizzo || row.via || ""}`);
    return anag ? `ANAG:${anag}` : "";
  }

  function summarizeData(operationalPlants = [], physicalPlants = [], workItems = []) {
    const groups = new Map();
    const plantIdToKey = new Map();

    const ensurePlant = (row = {}) => {
      const key = physicalIdentity(row, plantIdToKey);
      if (!key) return null;
      if (!groups.has(key)) groups.set(key, { key, plant: {}, work: [] });
      const group = groups.get(key);
      group.plant = { ...group.plant, ...row };
      [row.id, row.physicalPlantId, text(row.migrationSourceId).split("::")[0]].map(text).filter(Boolean)
        .forEach((id) => plantIdToKey.set(id, key));
      return group;
    };

    physicalPlants.filter(Boolean).forEach(ensurePlant);
    operationalPlants.filter(Boolean).forEach(ensurePlant);

    workItems.filter(Boolean).forEach((work) => {
      let key = physicalIdentity(work, plantIdToKey);
      if (!key) return;
      if (plantIdToKey.has(text(work.impiantoId))) key = plantIdToKey.get(text(work.impiantoId));
      if (!groups.has(key)) groups.set(key, { key, plant: { ...work }, work: [] });
      groups.get(key).work.push(work);
    });

    const items = [];
    let donePlants = 0;
    let doneWork = 0;
    let todoWork = 0;
    let subtotalCompleted = 0;

    for (const group of groups.values()) {
      group.work.forEach((work) => {
        if (isDone(work)) {
          doneWork += 1;
          const total = Number(work.totale);
          if (Number.isFinite(total)) subtotalCompleted += total;
        } else if (isTodo(work)) {
          todoWork += 1;
        }
      });

      const allWorkDone = group.work.length > 0 && group.work.every(isDone);
      const plantDone = group.work.length ? allWorkDone : isDone(group.plant);
      if (plantDone) donePlants += 1;

      let plantState = status(group.plant.statoGenerale || group.plant.stato) || "DA FARE";
      if (group.work.length) {
        plantState = allWorkDone ? "FATTO" : "DA FARE";
      }
      items.push({
        ...group.plant,
        done: plantDone,
        stato: plantState,
        statoGenerale: plantState,
        numeroLavorazioni: group.work.length,
        numeroLavorazioniFatte: group.work.filter(isDone).length,
        numeroLavorazioniDaFare: group.work.filter((item) => !isDone(item)).length,
        localFallback: false
      });
    }

    return {
      items,
      uniquePlants: groups.size,
      donePlants,
      todoPlants: Math.max(0, groups.size - donePlants),
      workRows: workItems.length,
      doneWork,
      todoWork,
      pendingWork: Math.max(0, workItems.length - doneWork),
      subtotalCompleted
    };
  }

  function summarizeManagementData(operationalPlants = [], physicalPlants = [], workItems = []) {
    const combined = summarizeData(operationalPlants, physicalPlants, workItems);
    if (!operationalPlants.length) return combined;

    const operational = summarizeData(operationalPlants, [], []);
    combined.uniquePlants = operational.uniquePlants;
    combined.donePlants = operational.donePlants;
    combined.todoPlants = operational.todoPlants;

    if (!workItems.length) {
      combined.workRows = operational.uniquePlants;
      combined.doneWork = operational.donePlants;
      combined.todoWork = operational.todoPlants;
      combined.pendingWork = operational.todoPlants;
    }
    return combined;
  }

  function renderDependentUI() {
    try { if (typeof recalculateCommessaWorkSummaries === "function") recalculateCommessaWorkSummaries(); } catch (_) {}
    try { if (typeof renderCommesseHomeList === "function") renderCommesseHomeList(); } catch (_) {}
    try { if (typeof renderCommesseManagementList === "function") renderCommesseManagementList(); } catch (_) {}
    try { if (typeof renderParentCommessaOverview === "function") renderParentCommessaOverview(); } catch (_) {}
    try { if (typeof updateCommessaDashboard === "function") updateCommessaDashboard(); } catch (_) {}
  }

  function removeSyntheticCommessaFromLocalState(canonical = findCanonicalModenaCommessa()) {
    if (!canonical?.id) return false;
    let changed = false;
    try {
      if (typeof unsubscribeCommessaStats !== "undefined" && unsubscribeCommessaStats?.has?.(SYNTHETIC_COMMESSA_ID)) {
        try { unsubscribeCommessaStats.get(SYNTHETIC_COMMESSA_ID)?.(); } catch (_) {}
        unsubscribeCommessaStats.delete(SYNTHETIC_COMMESSA_ID);
      }
      for (const map of [
        typeof commesseById !== "undefined" ? commesseById : null,
        typeof impiantiByCommessaId !== "undefined" ? impiantiByCommessaId : null,
        typeof commessaStatsById !== "undefined" ? commessaStatsById : null,
        typeof commessaWorkSummariesById !== "undefined" ? commessaWorkSummariesById : null,
        typeof commessaHoursById !== "undefined" ? commessaHoursById : null
      ]) {
        if (map?.delete) changed = map.delete(SYNTHETIC_COMMESSA_ID) || changed;
      }
      if (typeof selectedCommessaId !== "undefined" && selectedCommessaId === SYNTHETIC_COMMESSA_ID) selectedCommessaId = canonical.id;
      if (typeof managementCommessaId !== "undefined" && managementCommessaId === SYNTHETIC_COMMESSA_ID) managementCommessaId = canonical.id;
    } catch (_) {}
    if (changed) renderDependentUI();
    return changed;
  }

  async function archiveSyntheticDuplicate(canonical = findCanonicalModenaCommessa()) {
    if (!canonical?.id || activeCollectionName() !== "commesse") return false;
    removeSyntheticCommessaFromLocalState(canonical);
    if (!canWriteParentCommessa() || typeof db === "undefined" || !db || typeof auth === "undefined" || !auth.currentUser) return false;
    if (duplicateArchivePromise) return duplicateArchivePromise;

    duplicateArchivePromise = (async () => {
      try {
        const ref = db.collection("commesse").doc(SYNTHETIC_COMMESSA_ID);
        const snapshot = await ref.get();
        if (!snapshot.exists) return true;
        const current = snapshot.data() || {};
        if (current.attiva === false && current.hiddenFromHome === true && text(current.mergedIntoCommessaId) === text(canonical.id)) return true;
        const now = firebase.firestore.FieldValue.serverTimestamp();
        await ref.set({
          attiva: false,
          stato: "Archiviata",
          hiddenFromHome: true,
          mergedIntoCommessaId: canonical.id,
          mergedIntoCommessaCode: canonical.codice || canonical.code || CANONICAL_COMMESSA_CODE,
          mergeReason: "Duplicato tecnico sostituito dalla commessa INRETE MODENA canonica",
          mergedAt: now,
          updatedAt: now,
          updatedBy: auth.currentUser.uid
        }, { merge: true });
        console.info("[INRETE Modena] duplicato tecnico archiviato", { canonicalId: canonical.id });
        return true;
      } catch (error) {
        console.warn("[INRETE Modena] archiviazione duplicato non riuscita", error);
        return false;
      } finally {
        duplicateArchivePromise = null;
      }
    })();
    return duplicateArchivePromise;
  }

  function applyMetadataFallback(canonical) {
    if (!canonical?.id) return false;
    const total = Math.max(0, Number(canonical.impiantiCount || canonical.totalPlants || canonical.uniquePlantsCount || 0));
    if (!total) return false;
    const done = Math.min(total, Math.max(0, Number(canonical.impiantiFattiCount || canonical.donePlantsCount || 0)));
    try {
      if (typeof commessaStatsById !== "undefined" && commessaStatsById?.set) {
        const current = commessaStatsById.get(canonical.id) || {};
        if (!Number(current.total || 0)) commessaStatsById.set(canonical.id, { ...current, total, done });
      }
    } catch (_) {}
    renderDependentUI();
    return true;
  }

  function applySummary(canonical, summary) {
    if (!canonical?.id || !summary) return false;
    const commessaPatch = {
      impiantiCount: summary.uniquePlants,
      totalPlants: summary.uniquePlants,
      uniquePlantsCount: summary.uniquePlants,
      impiantiFattiCount: summary.donePlants,
      impiantiDaFareCount: summary.todoPlants,
      workItemsCount: summary.workRows,
      workItemsFattiCount: summary.doneWork,
      workItemsDaFareCount: summary.pendingWork,
      completedSubtotal: summary.subtotalCompleted,
      subtotalCompleted: summary.subtotalCompleted,
      summaryModelVersion: 3
    };

    try {
      if (typeof commesseById !== "undefined" && commesseById?.set) {
        const current = commesseById.get(canonical.id) || canonical;
        commesseById.set(canonical.id, { ...current, ...commessaPatch });
      }
      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.set) {
        impiantiByCommessaId.set(canonical.id, summary.items.map((item) => ({ ...item, commessaId: canonical.id })));
      }
      if (typeof commessaStatsById !== "undefined" && commessaStatsById?.set) {
        let stats = {};
        try {
          if (typeof calculateImpiantiStats === "function") stats = calculateImpiantiStats(summary.items) || {};
        } catch (_) {}
        commessaStatsById.set(canonical.id, { ...stats, total: summary.uniquePlants, done: summary.donePlants });
      }
      if (typeof selectedCommessaId !== "undefined" && selectedCommessaId === canonical.id && typeof currentImpianti !== "undefined") {
        currentImpianti = summary.items.map((item) => ({ ...item, commessaId: canonical.id }));
        try { if (typeof renderImpianti === "function") renderImpianti(); } catch (_) {}
        try { if (typeof renderMap === "function") renderMap(); } catch (_) {}
      }
    } catch (error) {
      console.warn("[INRETE Modena] applicazione riepilogo locale non riuscita", error);
    }

    renderDependentUI();
    applyManagementStats(canonical, summary);
    return true;
  }

  async function persistSummaryIfNeeded(canonical, summary) {
    if (!canonical?.id || !summary || !canWriteParentCommessa() || typeof db === "undefined" || !db || typeof auth === "undefined" || !auth.currentUser) return false;
    const fields = {
      impiantiCount: summary.uniquePlants,
      totalPlants: summary.uniquePlants,
      uniquePlantsCount: summary.uniquePlants,
      impiantiFattiCount: summary.donePlants,
      impiantiDaFareCount: summary.todoPlants,
      workItemsCount: summary.workRows,
      workItemsFattiCount: summary.doneWork,
      workItemsDaFareCount: summary.pendingWork
    };
    const changed = Object.entries(fields).some(([key, value]) => Number(canonical[key] ?? -1) !== Number(value));
    if (!changed) return false;
    try {
      const now = firebase.firestore.FieldValue.serverTimestamp();
      await db.collection("commesse").doc(canonical.id).set({
        ...fields,
        completedSubtotal: summary.subtotalCompleted,
        subtotalCompleted: summary.subtotalCompleted,
        summaryModelVersion: 3,
        summarySyncedAt: now,
        updatedAt: now,
        updatedBy: auth.currentUser.uid
      }, { merge: true });
      return true;
    } catch (error) {
      console.warn("[INRETE Modena] salvataggio riepilogo commessa non riuscito", error);
      return false;
    }
  }

  async function refreshCanonicalSummary(options = {}) {
    const canonical = findCanonicalModenaCommessa();
    if (!canonical?.id || activeCollectionName() !== "commesse" || typeof db === "undefined" || !db || typeof auth === "undefined" || !auth.currentUser) return false;
    const force = options.force === true;
    if (!force && lastSummary && Date.now() - lastSummaryAt < SUMMARY_REFRESH_MS) {
      applySummary(canonical, lastSummary);
      return lastSummary;
    }
    if (summaryPromise) return summaryPromise;

    summaryPromise = (async () => {
      try {
        const ref = db.collection("commesse").doc(canonical.id);
        const [operationalSnapshot, physicalSnapshot, workSnapshot] = await Promise.all([
          ref.collection("impianti").get(),
          ref.collection("impiantiFisici").get(),
          ref.collection("lavorazioni").get()
        ]);
        const operational = operationalSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        const physical = physicalSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        const work = workSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        const summary = summarizeData(operational, physical, work);
        lastSummary = summary;
        lastSummaryAt = Date.now();
        applySummary(canonical, summary);
        await persistSummaryIfNeeded(canonical, summary);
        return summary;
      } catch (error) {
        console.warn("[INRETE Modena] lettura riepilogo canonico non riuscita", error);
        return false;
      } finally {
        summaryPromise = null;
      }
    })();
    return summaryPromise;
  }

  function getManagementMetaCode() {
    const meta = document.querySelector?.("#impianti-management-meta");
    const value = text(meta?.textContent);
    const match = value.match(/(?:Cod\.?|Codice)\s*[:.]?\s*([A-Za-z0-9_-]+)/i);
    return normalizeCode(match?.[1] || "");
  }

  function isManagementForCommessa(commessa = {}) {
    if (!commessa?.id) return false;
    try {
      const activeId = typeof managementCommessaId !== "undefined" ? text(managementCommessaId) : "";
      if (activeId) return activeId === text(commessa.id);
    } catch (_) {}
    const targetCode = normalizeCode(commessa.codice || commessa.code);
    return Boolean(targetCode && getManagementMetaCode() === targetCode);
  }

  function findCurrentManagementCommessa() {
    try {
      const activeId = typeof managementCommessaId !== "undefined" ? text(managementCommessaId) : "";
      if (activeId && typeof commesseById !== "undefined" && commesseById?.get) {
        const current = commesseById.get(activeId);
        if (current) return { id: activeId, ...current };
      }
    } catch (_) {}

    const code = getManagementMetaCode();
    if (!code || typeof commesseById === "undefined" || !commesseById?.entries) return null;
    for (const [id, raw] of commesseById.entries()) {
      const commessa = { id, ...(raw || {}) };
      if (normalizeCode(commessa.codice || commessa.code) === code) return commessa;
    }
    return null;
  }

  function isManagementScreenVisible() {
    const screen = document.querySelector?.("#impianti-management-screen");
    if (!screen) return false;
    try {
      return !screen.classList?.contains?.("hidden");
    } catch (_) {
      return true;
    }
  }

  function renderManagementStatsForCommessa(commessa, summary) {
    if (!commessa?.id || !summary || !isManagementForCommessa(commessa)) return false;
    const target = document.querySelector?.("#impianti-management-stats");
    if (!target) return false;
    const money = new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" });
    const html = [
      `<span><b>${summary.uniquePlants}</b> impianti</span>`,
      `<span><b>${summary.workRows}</b> lavorazioni</span>`,
      `<span class="is-done"><b>${summary.doneWork}</b> lavorazioni fatte</span>`,
      `<span><b>${summary.pendingWork}</b> lavorazioni da fare</span>`,
      `<span><b>${money.format(summary.subtotalCompleted || 0)}</b> Subtotale lavorazioni fatte</span>`
    ].join("");
    if (target.innerHTML !== html) target.innerHTML = html;
    return true;
  }

  function applyManagementStats(canonical = findCanonicalModenaCommessa(), summary = lastSummary) {
    return renderManagementStatsForCommessa(canonical, summary);
  }

  async function refreshCurrentManagementSummary(options = {}) {
    const commessa = options.commessa || findCurrentManagementCommessa();
    if (!commessa?.id || typeof db === "undefined" || !db || typeof auth === "undefined" || !auth.currentUser) return false;
    const force = options.force === true;
    const cached = managementSummaryById.get(commessa.id);
    const cachedAt = Number(managementSummaryAtById.get(commessa.id) || 0);
    if (!force && cached && Date.now() - cachedAt < MANAGEMENT_DATA_REFRESH_MS) {
      renderManagementStatsForCommessa(commessa, cached);
      return cached;
    }
    if (managementSummaryPromiseById.has(commessa.id)) return managementSummaryPromiseById.get(commessa.id);

    const promise = (async () => {
      try {
        const ref = db.collection(activeCollectionName()).doc(commessa.id);
        const [operationalSnapshot, physicalSnapshot, workSnapshot] = await Promise.all([
          ref.collection("impianti").get(),
          ref.collection("impiantiFisici").get(),
          ref.collection("lavorazioni").get()
        ]);
        const operational = operationalSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        const physical = physicalSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        const work = workSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
        const summary = summarizeManagementData(operational, physical, work);
        managementSummaryById.set(commessa.id, summary);
        managementSummaryAtById.set(commessa.id, Date.now());
        renderManagementStatsForCommessa(commessa, summary);
        return summary;
      } catch (error) {
        console.warn("[Gestione commessa] riepilogo reale impianti non disponibile", { commessaId: commessa.id, error });
        return false;
      } finally {
        managementSummaryPromiseById.delete(commessa.id);
      }
    })();
    managementSummaryPromiseById.set(commessa.id, promise);
    return promise;
  }

  function ensureCurrentManagementSummary(options = {}) {
    if (!isManagementScreenVisible()) return false;
    const commessa = findCurrentManagementCommessa();
    if (!commessa?.id) return false;
    const cached = managementSummaryById.get(commessa.id);
    if (cached) renderManagementStatsForCommessa(commessa, cached);
    const cachedAt = Number(managementSummaryAtById.get(commessa.id) || 0);
    if ((options.force === true || !cached || Date.now() - cachedAt >= MANAGEMENT_DATA_REFRESH_MS)
      && (typeof document === "undefined" || !document.hidden)) {
      refreshCurrentManagementSummary({ commessa, force: options.force === true }).catch(() => {});
    }
    return true;
  }

  function ensureCanonicalState(options = {}) {
    const canonical = findCanonicalModenaCommessa();
    if (!canonical?.id) return false;
    removeSyntheticCommessaFromLocalState(canonical);
    applyMetadataFallback(canonical);
    archiveSyntheticDuplicate(canonical).catch(() => {});
    if (options.refresh !== false && (typeof document === "undefined" || !document.hidden)) {
      refreshCanonicalSummary({ force: options.force === true }).catch(() => {});
    }
    return true;
  }

  function installHistoricalCommesseResubscribe() {
    const GLOBAL = "HeraHistoricalCommesseResubscribe";
    if (window[GLOBAL]?.installed) return window[GLOBAL];
    const state = { installed: true, refreshed: false, attempts: 0, lastError: "" };
    let retryTimer = null;

    const refresh = () => {
      if (state.refreshed) return true;
      state.attempts += 1;
      try {
        const QueryPrototype = window.firebase?.firestore?.Query?.prototype;
        const restored = QueryPrototype?.__heraActiveCommesseOriginalOnSnapshot;
        const ready = typeof currentUser !== "undefined" && Boolean(currentUser)
          && typeof restored === "function" && QueryPrototype?.onSnapshot === restored
          && typeof stopCommesseSubscription === "function"
          && typeof subscribeCommesse === "function";
        if (!ready) {
          if (state.attempts < 40 && !retryTimer) retryTimer = setTimeout(() => { retryTimer = null; refresh(); }, 250);
          return false;
        }
        stopCommesseSubscription();
        const result = subscribeCommesse();
        if (result?.catch) result.catch((error) => { state.lastError = text(error?.message || error); });
        state.refreshed = true;
        return true;
      } catch (error) {
        state.lastError = text(error?.message || error);
        return false;
      }
    };

    window[GLOBAL] = { installed: true, refresh, getState: () => ({ ...state }) };
    setTimeout(refresh, 0);
    return window[GLOBAL];
  }

  window.createInreteModenaAugust2026 = async () => {
    ensureCanonicalState({ force: true });
    return Boolean(findCanonicalModenaCommessa());
  };
  window.refreshInreteModenaMixedWork = () => refreshCanonicalSummary({ force: true });
  window.INRETE_MODENA_AUGUST_2026 = {
    commessa: null,
    plants: [],
    ensureVisibleLocally: ensureCanonicalState,
    refreshWorkKinds: () => refreshCanonicalSummary({ force: true }),
    mergePlants: (...collections) => summarizeData(collections.flat().filter(Boolean), [], []).items,
    testing: {
      findCanonicalModenaCommessa,
      physicalIdentity,
      summarizeData,
      summarizeManagementData,
      removeSyntheticCommessaFromLocalState,
      applyMetadataFallback,
      getManagementMetaCode,
      isManagementForCommessa,
      findCurrentManagementCommessa,
      renderManagementStatsForCommessa,
      refreshCurrentManagementSummary,
      applyManagementStats
    }
  };

  window.HeraManagementCommessaSummary = {
    refresh: (commessa) => refreshCurrentManagementSummary({ commessa, force: true }),
    getCurrent: () => {
      const commessa = findCurrentManagementCommessa();
      return commessa?.id ? managementSummaryById.get(commessa.id) || null : null;
    }
  };

  installHistoricalCommesseResubscribe();
  setInterval(() => ensureCanonicalState(), LOCAL_GUARD_MS);
  setInterval(() => {
    applyManagementStats();
    ensureCurrentManagementSummary();
  }, MANAGEMENT_REFRESH_MS);

  try {
    if (typeof auth !== "undefined" && auth?.onAuthStateChanged) {
      auth.onAuthStateChanged((user) => {
        if (user) {
          setTimeout(() => ensureCanonicalState({ force: true }), 300);
          setTimeout(() => ensureCurrentManagementSummary({ force: true }), 500);
        }
      });
    }
  } catch (_) {}

  window.addEventListener("load", () => {
    setTimeout(() => ensureCanonicalState({ force: true }), 300);
    setTimeout(() => ensureCanonicalState({ force: true }), 1800);
    setTimeout(() => ensureCurrentManagementSummary({ force: true }), 1200);
  });
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      ensureCanonicalState({ force: true });
      ensureCurrentManagementSummary({ force: true });
    }
  });
  setTimeout(() => ensureCanonicalState(), 100);
})();

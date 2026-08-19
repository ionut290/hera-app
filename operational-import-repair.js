/* Crea INRETE MODENA - AGOSTO 2026 senza sovrascrivere stati operativi già salvati. */
(() => {
  "use strict";

  const COMMESSA_ID = "inrete_modena_agosto_2026";
  const DATASET_VERSION = "2026-08-19-ago30-ui-fallback-v4-safe-status";
  const RETRY_MS = 2000;
  const MAX_ATTEMPTS = 90;
  const MIXED_REFRESH_MS = 30000;
  let attempts = 0;
  let running = false;
  let mixedRefreshRunning = false;
  let workItemsCache = [];

  const RAW = Object.freeze([
    ["3430707","REMI CASTELFRANCO","CASTELFRANCO EMILIA","Via Loda, 28",44.58931,11.04352],
    ["3430051","REMI S.CESARIO","S.CESARIO SUL PANARO","Via della Cartiera, 33 (dietro al Cimitero)",44.56565,11.03273],
    ["3472968","REMI SAVIGNANO","SAVIGNANO SUL PANARO","Via S. Anna, 10",44.48468,11.03210],
    ["3471231","REMI SPEZZANO","FIORANO MODENESE","Via Crociale, 7",44.53632,10.83482],
    ["3470478","REMI FORMIGINE","FORMIGINE","Via Grandi",44.56882,10.84307],
    ["3475570","REMI MAGRETA","SASSUOLO","Via Secchia, 32",44.59869,10.79012],
    ["3471127","REMI CASINALBO","FORMIGINE","Via Sant Ambrogio",44.59549,10.84845],
    ["3471152","REMI MARANELLO","MARANELLO","Via Claudia",44.52638,10.85457],
    ["3470982","REMI POZZA","MARANELLO","Via Vandelli",44.53058,10.89367],
    ["3471269","REMI BRAIDA","SASSUOLO","Via S.Pietro/via San Bernardo",44.54703,10.79935],
    ["area adiacente alla REMI","STOCCAGGIO TUBI BRAIDA","SASSUOLO","Via S.Pietro/via San Bernardo",44.54703,10.79935],
    ["3470559","REMI INDIPENDENZA","SASSUOLO","Via Indipendenza",44.53781,10.77355],
    ["3471102","REMI PONTE FOSSA","SASSUOLO","Via Valle d'Aosta",44.56558,10.80002],
    ["3471279","REMI UBERSETTO","FORMIGINE","Via dei Prati",44.55092,10.86203],
    ["3429985","REMI CASTELNUOVO","CASTELNUOVO RANGONE","Via Gualinga, 23",44.53866,10.94932],
    ["3430270","REMI CA' DI SOLA","CASTELVETRO","Via Per Modena",44.53481,10.95387],
    ["3430320","REMI SOLIGNANO","CASTELVETRO","Via Montanara, 4",44.53274,10.91688],
    ["3425784","REMI S.CLEMENTE","MODENA","strada di S. Clemente, 11",44.70429,10.99211],
    ["3426714","REMI S.CLEMENTE","MODENA","strada di S. Clemente, 11",44.61267,10.97673],
    ["3430248","REMI SUD","MODENA","Via Cadiane, 255",44.60603,10.91114],
    ["3430604","REMI CASTELLARO","SPILAMBERTO","Via Castellaro, 13",44.54071,11.02894],
    ["3430547","REMI S.VITO","SPILAMBERTO","Via San Vito",44.53941,11.01387],
    ["3430221","REMI VIGNOLA","VIGNOLA","Via Doccia",44.48418,11.02692],
    ["3476117","REMI BERZIGALA","SERRAMAZZONI","Via Giardini (Loc. Ca Ambero)",44.39255,10.80896],
    ["3473527","REMI PAVULLO CASTELLO","PAVULLO NEL FRIGNANO","Via Montecuccolo",44.32698,10.83440],
    ["3473528","REMI S.ANTONIO","PAVULLO NEL FRIGNANO","Via Pico/Guicciardini",44.36767,10.83285],
    ["3471537","REMI S. DALMAZIO","SERRAMAZZONI","Via Per Marano",44.42214,10.85448],
    ["3475075","REMI MONTECENERE","PAVULLO NEL FRIGNANO","Via Bellini",44.31195,10.78122],
    ["3426336","GRMI UMF07 _ EXPORT CERAM","MONTEFIORINO","VIA LA PIANA, 6",44.37793,10.61998],
    ["3426335","GRMI UMF08 _ EXPORT CERAM","MONTEFIORINO","VIA LA PIANA, 6",44.37797,10.62001]
  ]);

  const text = (value) => String(value ?? "").trim();
  const norm = (value) => text(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "");
  const present = (value) => value !== undefined && value !== null && value !== "";
  const plantId = (idSap, name, index) => /^\d+$/.test(String(idSap))
    ? `sap_${idSap}`
    : `modena_${String(name || index + 1).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "")}`;

  function baseSourceId(value) {
    let raw = text(value);
    for (const separator of ["::", "__"]) {
      const index = raw.indexOf(separator);
      if (index > 0) raw = raw.slice(0, index);
    }
    return raw;
  }

  function sourceIdsFor(item = {}) {
    return [...new Set([
      item.id,
      item.physicalPlantId,
      item.impiantoId,
      item.migrationSourceId,
      ...(Array.isArray(item.sourceIds) ? item.sourceIds : [])
    ].map(text).filter(Boolean))];
  }

  function identity(item = {}) {
    const sap = norm(item.idSap || item.idSAP || item["ID SAP"] || item.codiceSap);
    if (sap) return `sap:${sap}`;
    const name = item.denominazione || item.nome || item.name || "";
    const comune = item.comune || "";
    const address = item.indirizzo || item.via || item.descrizioneVia || "";
    const anagrafica = norm(`${name}|${comune}|${address}`);
    if (anagrafica) return `anag:${anagrafica}`;
    const sourceId = baseSourceId(sourceIdsFor(item)[0]);
    return sourceId ? `id:${norm(sourceId)}` : "";
  }

  function detectKinds(item = {}) {
    const raw = [item.tipo, item.categoria, item.tipologia, item.tipologiaLavorazione, item.tipologiaIntervento, item.descrizione, item.note, item.attivitaLabel]
      .map(text).join(" ").toLocaleUpperCase("it-IT");
    const kinds = new Set();
    if (raw.includes("STRAORD")) kinds.add("STRAORDINARIO");
    if (raw.replace(/STRAORD\w*/g, " ").includes("ORDIN")) kinds.add("ORDINARIO");
    return kinds;
  }

  const PLANTS = RAW.map((row, index) => {
    const [idSap, denominazione, comune, indirizzo, lat, lng] = row;
    const id = plantId(idSap, denominazione, index);
    return {
      id,
      commessaId: COMMESSA_ID,
      physicalPlantId: id,
      migrationSourceId: id,
      numeroProgressivo: index + 1,
      numeroProgressivoImpianto: index + 1,
      distretto: "Modena",
      idSap,
      denominazione,
      nome: denominazione,
      comune,
      indirizzo,
      via: indirizzo,
      gpsY: lat,
      gpsX: lng,
      latitudine: lat,
      longitudine: lng,
      stato: "DA FARE",
      statoGenerale: "DA FARE",
      done: false,
      tipo: "ORDINARIO",
      tipologiaIntervento: "ORDINARIO",
      attivitaLabel: "ORDINARIO",
      sourceDataset: "06-Attività Sfalci - Ago-2026 - MO.xlsx",
      sourceDatasetVersion: DATASET_VERSION
    };
  });

  const COMMESSA = {
    id: COMMESSA_ID,
    nome: "INRETE MODENA - AGOSTO 2026",
    name: "INRETE MODENA - AGOSTO 2026",
    codice: "INRETE-MO-AGO-2026",
    code: "INRETE-MO-AGO-2026",
    cliente: "INRETE",
    categoria: "INRETE GAS MODENA",
    area: "MODENA",
    stato: "Attiva",
    attiva: true,
    parentCommessaId: null,
    excelModelVersion: 2,
    priceListVersion: 2,
    workItemsModelVersion: 2,
    percentualeRibassoGenerale: 0.01,
    nextImpiantoNumber: 31,
    datasetVersion: DATASET_VERSION,
    sourceDataset: "06-Attività Sfalci - Ago-2026 - MO.xlsx",
    impiantiCount: PLANTS.length,
    totalPlants: PLANTS.length,
    impiantiFattiCount: 0,
    impiantiDaFareCount: PLANTS.length,
    localFallback: true
  };

  const collectionName = () => "commesse";

  function mergePreferred(base, incoming) {
    const out = { ...(base || {}) };
    Object.entries(incoming || {}).forEach(([field, value]) => {
      if (!present(value)) return;
      if (["gpsY", "gpsX", "latitudine", "longitudine", "latitude", "longitude"].includes(field) && Number(value) === 0) return;
      out[field] = value;
    });
    return out;
  }

  function mergePlants(...collections) {
    const groups = new Map();
    const sourceToGroup = new Map();

    const registerSources = (groupKey, group, item) => {
      sourceIdsFor(item).forEach((sourceId) => {
        group.sourceIds.add(sourceId);
        const baseId = baseSourceId(sourceId);
        if (baseId) {
          group.sourceIds.add(baseId);
          sourceToGroup.set(norm(baseId), groupKey);
        }
        sourceToGroup.set(norm(sourceId), groupKey);
      });
    };

    const pushPlant = (item) => {
      if (!item) return;
      const key = identity(item);
      if (!key) return;
      let group = groups.get(key);
      if (!group) {
        group = { item: { ...item }, kinds: new Set(), sourceIds: new Set() };
        groups.set(key, group);
      } else {
        group.item = mergePreferred(group.item, item);
      }
      detectKinds(item).forEach((kind) => group.kinds.add(kind));
      registerSources(key, group, item);
    };

    collections.flat().filter(Boolean).forEach(pushPlant);

    workItemsCache.forEach((item) => {
      const sourceKeys = sourceIdsFor(item).flatMap((sourceId) => [norm(sourceId), norm(baseSourceId(sourceId))]).filter(Boolean);
      let groupKey = sourceKeys.map((key) => sourceToGroup.get(key)).find(Boolean) || "";
      if (!groupKey) {
        const fallbackIdentity = identity(item);
        if (fallbackIdentity && groups.has(fallbackIdentity)) groupKey = fallbackIdentity;
      }
      const group = groupKey ? groups.get(groupKey) : null;
      if (!group) return;
      detectKinds(item).forEach((kind) => group.kinds.add(kind));
      registerSources(groupKey, group, item);
    });

    return [...groups.values()].map((group) => {
      const out = { ...group.item, sourceIds: [...group.sourceIds] };
      const hasOrdinary = group.kinds.has("ORDINARIO");
      const hasExtraordinary = group.kinds.has("STRAORDINARIO");
      if (hasOrdinary && hasExtraordinary) {
        out.tipo = "ORDINARIO E STRAORDINARIO";
        out.tipologia = "ORDINARIO E STRAORDINARIO";
        out.tipologiaIntervento = "ORDINARIO E STRAORDINARIO";
        out.tipologiaLavorazione = "ORDINARIO E STRAORDINARIO";
        out.attivitaLabel = "ORDINARIO E STRAORDINARIO";
        out.isMixedOrdinaryExtraordinary = true;
        out.markerColor = "#f4c542";
        out.markerFillColor = "#f4c542";
        out.markerClass = "impianto-marker-mixed-yellow";
        out.markerTone = "yellow";
      } else {
        if (out.markerClass === "impianto-marker-mixed-yellow") {
          delete out.markerColor;
          delete out.markerFillColor;
          delete out.markerClass;
          delete out.markerTone;
        }
        out.isMixedOrdinaryExtraordinary = false;
        if (hasExtraordinary) {
          out.tipo = "STRAORDINARIO";
          out.tipologiaIntervento = "STRAORDINARIO";
          out.attivitaLabel = "STRAORDINARIO";
        } else if (hasOrdinary) {
          out.tipo = "ORDINARIO";
          out.tipologiaIntervento = "ORDINARIO";
          out.attivitaLabel = "ORDINARIO";
        }
      }
      return out;
    }).sort((a, b) => (Number(a.numeroProgressivo || a.numeroProgressivoImpianto) || 9999) - (Number(b.numeroProgressivo || b.numeroProgressivoImpianto) || 9999));
  }

  function signature(items) {
    return (Array.isArray(items) ? items : []).map((item) => [
      identity(item),
      item.attivitaLabel,
      item.statoGenerale || item.stato,
      item.done ? 1 : 0,
      item.numeroLavorazioniFatte,
      item.numeroLavorazioniDaFare
    ].join("|")).join(";");
  }

  function applyMergedVisiblePlants() {
    let cached = [];
    try {
      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.get) cached = impiantiByCommessaId.get(COMMESSA_ID) || [];
    } catch (_) {}
    let current = [];
    try {
      if (typeof selectedCommessaId !== "undefined" && selectedCommessaId === COMMESSA_ID && typeof currentImpianti !== "undefined" && Array.isArray(currentImpianti)) current = currentImpianti;
    } catch (_) {}

    const merged = mergePlants(PLANTS, cached, current);
    try {
      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.set && signature(cached) !== signature(merged)) {
        impiantiByCommessaId.set(COMMESSA_ID, merged.map((item) => ({ ...item })));
      }
    } catch (_) {}

    try {
      if (typeof selectedCommessaId !== "undefined" && selectedCommessaId === COMMESSA_ID && typeof currentImpianti !== "undefined" && signature(current) !== signature(merged)) {
        currentImpianti = merged.map((item) => ({ ...item }));
        if (typeof renderImpianti === "function") renderImpianti();
        if (typeof renderMap === "function") renderMap();
        if (typeof renderHeaderActivitySummary === "function") renderHeaderActivitySummary();
        if (typeof updateCommessaDashboard === "function") updateCommessaDashboard();
      }
    } catch (error) {
      console.warn("[INRETE Modena merge ordinario/straordinario UI]", error);
    }
    return merged;
  }

  async function refreshWorkKinds() {
    if (mixedRefreshRunning || typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser) return false;
    mixedRefreshRunning = true;
    try {
      const ref = db.collection(collectionName()).doc(COMMESSA_ID);
      const snapshot = await ref.collection("lavorazioni").get();
      workItemsCache = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      applyMergedVisiblePlants();
      return true;
    } catch (error) {
      console.warn("[INRETE Modena] lettura lavorazioni ordinario/straordinario non riuscita", error);
      return false;
    } finally {
      mixedRefreshRunning = false;
    }
  }

  function ensureVisibleLocally() {
    try {
      let changed = false;
      if (typeof commesseById !== "undefined" && commesseById?.set && !commesseById.has(COMMESSA_ID)) {
        commesseById.set(COMMESSA_ID, { ...COMMESSA });
        changed = true;
      }
      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.set) {
        const cached = impiantiByCommessaId.get(COMMESSA_ID);
        if (!Array.isArray(cached) || cached.length === 0) {
          impiantiByCommessaId.set(COMMESSA_ID, PLANTS.map((item) => ({ ...item })));
          changed = true;
        }
      }
      if (changed) {
        if (typeof renderCommesseHomeList === "function") renderCommesseHomeList();
        if (typeof renderCommesseManagementList === "function") renderCommesseManagementList();
        if (typeof refreshCommesseDependentUI === "function") refreshCommesseDependentUI(false);
      }
      if (typeof selectedCommessaId !== "undefined" && selectedCommessaId === COMMESSA_ID && typeof currentImpianti !== "undefined") {
        if (!Array.isArray(currentImpianti) || currentImpianti.length === 0) {
          currentImpianti = PLANTS.map((item) => ({ ...item }));
        }
        applyMergedVisiblePlants();
      }
    } catch (error) {
      console.warn("[INRETE Modena fallback UI]", error);
    }
  }

  function canWriteCatalog() {
    try {
      return typeof canManageData === "function" && canManageData() === true;
    } catch (_) {
      return false;
    }
  }

  function commessaPayload(existingData, isNew) {
    const now = firebase.firestore.FieldValue.serverTimestamp();
    if (isNew) {
      return {
        ...COMMESSA,
        localFallback: false,
        creatoDa: auth.currentUser.email || "",
        createdBy: auth.currentUser.uid,
        createdAt: now,
        updatedAt: now,
        operationalModelSyncedAt: now
      };
    }
    return {
      datasetVersion: DATASET_VERSION,
      sourceDataset: COMMESSA.sourceDataset,
      localFallback: false,
      updatedAt: now,
      operationalModelSyncedAt: now,
      operationalCatalogUpdatedBy: auth.currentUser.uid,
      operationalCatalogPreviousVersion: text(existingData.datasetVersion)
    };
  }

  function plantPayload(plant, existingData, exists) {
    const now = firebase.firestore.FieldValue.serverTimestamp();
    const payload = {
      id: plant.id,
      commessaId: COMMESSA_ID,
      physicalPlantId: plant.physicalPlantId,
      migrationSourceId: plant.migrationSourceId,
      numeroProgressivo: plant.numeroProgressivo,
      numeroProgressivoImpianto: plant.numeroProgressivoImpianto,
      distretto: plant.distretto,
      idSap: plant.idSap,
      denominazione: plant.denominazione,
      nome: plant.nome,
      comune: plant.comune,
      indirizzo: plant.indirizzo,
      via: plant.via,
      gpsY: plant.gpsY,
      gpsX: plant.gpsX,
      latitudine: plant.latitudine,
      longitudine: plant.longitudine,
      sourceDataset: plant.sourceDataset,
      sourceDatasetVersion: DATASET_VERSION,
      updatedAt: now
    };
    if (!exists) payload.createdAt = now;
    if (!present(existingData?.stato)) payload.stato = "DA FARE";
    if (!present(existingData?.statoGenerale)) payload.statoGenerale = "DA FARE";
    if (existingData?.done === undefined || existingData?.done === null) payload.done = false;
    if (!present(existingData?.tipo)) payload.tipo = "ORDINARIO";
    if (!present(existingData?.tipologiaIntervento)) payload.tipologiaIntervento = "ORDINARIO";
    if (!present(existingData?.attivitaLabel)) payload.attivitaLabel = "ORDINARIO";
    return payload;
  }

  async function createInFirestore(force = false) {
    if (running || typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser || !canWriteCatalog()) return false;
    running = true;
    try {
      const ref = db.collection(collectionName()).doc(COMMESSA_ID);
      const existing = await ref.get();
      const existingData = existing.exists ? (existing.data() || {}) : {};
      if (!force && existing.exists && text(existingData.datasetVersion) === DATASET_VERSION) return true;

      const [operationalSnapshot, physicalSnapshot] = await Promise.all([
        ref.collection("impianti").get(),
        ref.collection("impiantiFisici").get()
      ]);
      const operationalById = new Map(operationalSnapshot.docs.map((doc) => [doc.id, doc.data() || {}]));
      const physicalById = new Map(physicalSnapshot.docs.map((doc) => [doc.id, doc.data() || {}]));

      await ref.set(commessaPayload(existingData, !existing.exists), { merge: true });

      const batch = db.batch();
      PLANTS.forEach((plant) => {
        const operationalExisting = operationalById.get(plant.id);
        const physicalExisting = physicalById.get(plant.id);
        batch.set(
          ref.collection("impianti").doc(plant.id),
          plantPayload(plant, operationalExisting, operationalById.has(plant.id)),
          { merge: true }
        );
        batch.set(
          ref.collection("impiantiFisici").doc(plant.id),
          plantPayload(plant, physicalExisting, physicalById.has(plant.id)),
          { merge: true }
        );
      });
      await batch.commit();
      console.info("[INRETE Modena] catalogo sincronizzato senza sovrascrivere gli stati", { count: PLANTS.length });
      return true;
    } finally {
      running = false;
    }
  }

  async function tryRun() {
    ensureVisibleLocally();
    attempts += 1;
    try {
      if (typeof auth === "undefined" || typeof db === "undefined" || !auth.currentUser) throw new Error("Firebase non pronto");
      if (!canWriteCatalog()) return;
      if (await createInFirestore(false)) {
        await refreshWorkKinds();
        return;
      }
    } catch (error) {
      console.warn("[INRETE Modena] sincronizzazione catalogo non riuscita, resta attivo il fallback UI", attempts, error?.message || error);
    }
    if (attempts < MAX_ATTEMPTS) setTimeout(tryRun, RETRY_MS);
  }

  window.createInreteModenaAugust2026 = (options = {}) => createInFirestore(options.force === true);
  window.refreshInreteModenaMixedWork = refreshWorkKinds;
  window.INRETE_MODENA_AUGUST_2026 = {
    commessa: COMMESSA,
    plants: PLANTS,
    ensureVisibleLocally,
    refreshWorkKinds,
    mergePlants,
    detectKinds,
    identity
  };

  setInterval(ensureVisibleLocally, 1500);
  setInterval(() => {
    try {
      if (typeof selectedCommessaId !== "undefined" && selectedCommessaId === COMMESSA_ID) refreshWorkKinds();
    } catch (_) {}
  }, MIXED_REFRESH_MS);
  if (typeof auth !== "undefined" && auth?.onAuthStateChanged) auth.onAuthStateChanged((user) => { if (user) setTimeout(tryRun, 250); });
  window.addEventListener("load", () => { ensureVisibleLocally(); setTimeout(tryRun, 500); setTimeout(refreshWorkKinds, 1200); });
  document.addEventListener("visibilitychange", () => { if (!document.hidden) { ensureVisibleLocally(); refreshWorkKinds(); } });
  setTimeout(ensureVisibleLocally, 100);
  setTimeout(tryRun, 1000);
})();

/* Protegge commesse e impianti dalle sincronizzazioni Google Sheet distruttive. */
(() => {
  "use strict";

  const GLOBAL = "HeraGoogleSheetImportSafetyGuard";
  if (window[GLOBAL]?.installed) return;

  const state = {
    installed: false,
    guardedBatches: 0,
    blockedDeletes: 0,
    protectedStatusWrites: 0,
    lastProtectedPath: ""
  };
  const OPERATIONAL_FIELDS = [
    "stato", "statoGenerale", "done", "doneAt",
    "dataEsecuzione", "oraEsecuzione",
    "operatoreNome", "operatore", "operatoreUid", "operatoreEmail"
  ];
  const EXECUTION_FIELDS = ["dataEsecuzione", "oraEsecuzione", "operatoreNome", "operatore", "operatoreUid", "operatoreEmail"];
  let attempts = 0;

  const refPath = (ref) => String(ref?.path || ref?._key?.path?.canonicalString?.() || "");
  const isProtectedPath = (path) => /(?:^|\/)(?:commesse|globalCommesse)\/[^/]+\/(?:impianti|impiantiFisici|lavorazioni)\/[^/]+$/.test(path);
  const isGoogleSheetStack = (stack) => /google-sheet-two-way-sync\.js|applyRowsToFirestore|commitOperations/i.test(String(stack || ""));
  const isCompletedPayload = (data) => {
    const status = String(data?.stato || data?.statoGenerale || "").trim().toLocaleUpperCase("it-IT");
    return data?.done === true || ["FATTO", "DONE", "COMPLETATO"].includes(status);
  };

  function sanitizeSheetPayload(ref, data) {
    if (!data || typeof data !== "object") return data;
    const path = refPath(ref);
    if (!isProtectedPath(path)) return data;
    const safe = { ...data };
    if (!isCompletedPayload(data)) {
      OPERATIONAL_FIELDS.forEach((field) => delete safe[field]);
      state.protectedStatusWrites += 1;
      state.lastProtectedPath = path;
      return safe;
    }
    EXECUTION_FIELDS.forEach((field) => {
      if (safe[field] === undefined || safe[field] === null || String(safe[field]).trim() === "") delete safe[field];
    });
    return safe;
  }

  function install() {
    let runtimeDb = null;
    try { runtimeDb = typeof db !== "undefined" ? db : window.db; } catch (_) { runtimeDb = window.db; }
    if (!runtimeDb || typeof runtimeDb.batch !== "function") return false;
    if (runtimeDb.batch.__heraGoogleSheetImportSafetyGuard) {
      state.installed = true;
      return true;
    }

    const originalBatchFactory = runtimeDb.batch.bind(runtimeDb);
    const guardedBatchFactory = function guardedFirestoreBatch() {
      const originalBatch = originalBatchFactory();
      const creationStack = String(new Error().stack || "");
      let sheetManagedBatch = isGoogleSheetStack(creationStack);
      let proxy = null;

      proxy = new Proxy(originalBatch, {
        get(target, property, receiver) {
          if (property === "set") {
            return (ref, data, options) => {
              if (!sheetManagedBatch && (data?.sheetManaged === true || data?.sheetSyncSourceId)) {
                sheetManagedBatch = true;
                state.guardedBatches += 1;
              }
              const nextData = sheetManagedBatch ? sanitizeSheetPayload(ref, data) : data;
              target.set(ref, nextData, options);
              return proxy;
            };
          }
          if (property === "delete") {
            return (ref) => {
              const path = refPath(ref);
              if ((sheetManagedBatch || isGoogleSheetStack(new Error().stack)) && isProtectedPath(path)) {
                state.blockedDeletes += 1;
                state.lastProtectedPath = path;
                console.warn("[Google Sheet Safety] cancellazione bloccata", path);
                return proxy;
              }
              target.delete(ref);
              return proxy;
            };
          }
          const value = Reflect.get(target, property, receiver);
          return typeof value === "function" ? value.bind(target) : value;
        }
      });
      if (sheetManagedBatch) state.guardedBatches += 1;
      return proxy;
    };

    Object.defineProperty(guardedBatchFactory, "__heraGoogleSheetImportSafetyGuard", { value: true });
    Object.defineProperty(guardedBatchFactory, "__heraGoogleSheetImportSafetyOriginal", { value: originalBatchFactory });
    runtimeDb.batch = guardedBatchFactory;
    state.installed = true;
    return true;
  }

  window[GLOBAL] = {
    installed: false,
    version: "1.0.0",
    mode: "non-destructive-sheet-import",
    install,
    sanitizeSheetPayload,
    isProtectedPath,
    getState: () => ({ ...state })
  };

  const retry = () => {
    attempts += 1;
    if (install() || attempts >= 80) {
      window[GLOBAL].installed = state.installed;
      return;
    }
    setTimeout(retry, 250);
  };
  retry();
})();

/* Mantiene INRETE MODENA - AGOSTO 2026 visibile e sincronizzata senza sovrascrivere gli stati operativi. */
(() => {
  "use strict";

  const COMMESSA_ID = "inrete_modena_agosto_2026";
  const DATASET_VERSION = "2026-08-19-commesse-impianti-integrity-v4";
  const RETRY_MS = 2500;
  const MAX_ATTEMPTS = 60;
  const MIXED_REFRESH_MS = 60000;
  const SOURCE_DATASET = "06-Attività Sfalci - Ago-2026 - MO.xlsx";
  const MIXED_MARKER_COLOR = "#f4c542";
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
  const norm = (value) => text(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
  const hasValue = (value) => value !== undefined && value !== null && !(typeof value === "string" && value.trim() === "");
  const plantId = (idSap, name, index) => /^\d+$/.test(String(idSap))
    ? `sap_${idSap}`
    : `modena_${String(name || index + 1).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "")}`;

  function sapAlias(item = {}) {
    const sap = norm(item.idSap || item.idSAP || item["ID SAP"] || item.codiceSap);
    return sap ? `sap:${sap}` : "";
  }

  function anagAlias(item = {}) {
    const name = item.denominazione || item.nome || item.name || "";
    const comune = item.comune || "";
    const address = item.indirizzo || item.via || item.descrizioneVia || "";
    const normalized = norm(`${name}|${comune}|${address}`);
    return normalized ? `anag:${normalized}` : "";
  }

  function baseMigrationSourceId(value) {
    const raw = text(value);
    return raw ? raw.split("::")[0].trim() : "";
  }

  function extractPlantIds(item = {}, includeDocumentId = true) {
    const ids = new Set();
    const add = (value) => {
      const normalized = text(value);
      if (normalized) ids.add(normalized);
    };
    if (includeDocumentId) add(item.id);
    add(item.physicalPlantId);
    add(item.impiantoId);
    add(baseMigrationSourceId(item.migrationSourceId));
    if (Array.isArray(item.sourceIds)) item.sourceIds.forEach(add);
    return [...ids];
  }

  function identity(item = {}) {
    const sap = sapAlias(item);
    if (sap) return sap;
    const ids = extractPlantIds(item);
    if (ids.length) return `id:${norm(ids[0])}`;
    return anagAlias(item);
  }

  function plantLookupAliases(item = {}) {
    const aliases = [];
    const sap = sapAlias(item);
    if (sap) aliases.push(sap);
    extractPlantIds(item).forEach((id) => aliases.push(`id:${norm(id)}`));
    if (!aliases.length) {
      const anag = anagAlias(item);
      if (anag) aliases.push(anag);
    }
    return [...new Set(aliases.filter(Boolean))];
  }

  function workItemPlantAliases(item = {}) {
    const aliases = [];
    [item.physicalPlantId, item.impiantoId, baseMigrationSourceId(item.migrationSourceId)]
      .map(text)
      .filter(Boolean)
      .forEach((id) => aliases.push(`id:${norm(id)}`));
    const sap = sapAlias(item);
    if (sap) aliases.push(sap);
    if (!aliases.length) {
      const anag = anagAlias(item);
      if (anag) aliases.push(anag);
    }
    return [...new Set(aliases.filter(Boolean))];
  }

  function detectKinds(item = {}) {
    const raw = [
      item.tipo,
      item.categoria,
      item.tipologia,
      item.tipologiaLavorazione,
      item.tipologiaIntervento,
      item.descrizione,
      item.note,
      item.attivitaLabel
    ].map(text).join(" ").toLocaleUpperCase("it-IT");

    const kinds = new Set();
    if (!raw) return kinds;

    const hasExtraordinary = /\bSTRAORD(?:INARI[OAIE]?)?\b/.test(raw);
    const withoutExtraordinary = raw.replace(/\bSTRAORD(?:INARI[OAIE]?)?\b/g, " ");
    const hasOrdinary = /\bORDINARI[OAIE]?\b/.test(withoutExtraordinary);

    if (hasExtraordinary) kinds.add("STRAORDINARIO");
    if (hasOrdinary) kinds.add("ORDINARIO");
    return kinds;
  }

  function parseCoordinate(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    const normalized = text(value).replace(",", ".");
    if (!normalized) return null;
    const number = Number(normalized);
    return Number.isFinite(number) ? number : null;
  }

  function isValidCoordinatePair(lat, lng) {
    return Number.isFinite(lat) && Number.isFinite(lng)
      && lat >= -90 && lat <= 90
      && lng >= -180 && lng <= 180;
  }

  function isModenaCoordinatePair(lat, lng) {
    return isValidCoordinatePair(lat, lng)
      && lat >= 43.5 && lat <= 45.2
      && lng >= 9.5 && lng <= 12.5;
  }

  function normalizeCoordinatePair(first, second) {
    const a = parseCoordinate(first);
    const b = parseCoordinate(second);
    if (!isValidCoordinatePair(a, b)) return null;
    if (isModenaCoordinatePair(a, b)) return { lat: a, lng: b };
    if (isModenaCoordinatePair(b, a)) return { lat: b, lng: a };
    return { lat: a, lng: b };
  }

  function coordinatePairFrom(item = {}) {
    const candidates = [
      [item.latitudine, item.longitudine],
      [item.latitude, item.longitude],
      [item.lat, item.lng],
      [item.gpsY, item.gpsX],
      [item.coordinateGPSY, item.coordinateGPSX],
      [item["Coordinate GPS(Y)"], item["Coordinate GPS(X)"]]
    ];

    for (const [lat, lng] of candidates) {
      const pair = normalizeCoordinatePair(lat, lng);
      if (pair && isModenaCoordinatePair(pair.lat, pair.lng)) return pair;
    }

    const coordinateText = text(item.coordinate || item.coordinates || item.coordinateGps || item.coordinateGPS);
    if (coordinateText) {
      const values = coordinateText.match(/-?\d+(?:[.,]\d+)?/g) || [];
      if (values.length >= 2) {
        const pair = normalizeCoordinatePair(values[0], values[1]);
        if (pair && isModenaCoordinatePair(pair.lat, pair.lng)) return pair;
      }
    }

    return null;
  }

  function applyCoordinatePair(target, pair) {
    if (!pair) return target;
    target.gpsY = pair.lat;
    target.gpsX = pair.lng;
    target.latitudine = pair.lat;
    target.longitudine = pair.lng;
    target.latitude = pair.lat;
    target.longitude = pair.lng;
    target.lat = pair.lat;
    target.lng = pair.lng;
    return target;
  }

  function normalizeWorkStatus(item = {}) {
    return text(item.stato || item.status || item.statoLavorazione).toLocaleUpperCase("it-IT");
  }

  function isDoneStatus(value) {
    const status = text(value).toLocaleUpperCase("it-IT");
    return /(^|\s)(FATTO|COMPLETATO|COMPLETATA|CHIUSO|CHIUSA|ESEGUITO|ESEGUITA)(\s|$)/.test(status);
  }

  function isPartialStatus(value) {
    const status = text(value).toLocaleUpperCase("it-IT");
    return status.includes("PARZIAL") || status.includes("IN LAVORAZIONE") || status.includes("IN CORSO");
  }

  function deriveWorkAggregate(items = []) {
    const rows = Array.isArray(items) ? items.filter(Boolean) : [];
    let doneCount = 0;
    let startedCount = 0;
    rows.forEach((item) => {
      const status = normalizeWorkStatus(item);
      if (item.done === true || isDoneStatus(status)) {
        doneCount += 1;
        startedCount += 1;
      } else if (isPartialStatus(status)) {
        startedCount += 1;
      }
    });
    return {
      total: rows.length,
      done: doneCount,
      pending: Math.max(0, rows.length - doneCount),
      allDone: rows.length > 0 && doneCount === rows.length,
      partial: rows.length > 0 && doneCount < rows.length && startedCount > 0
    };
  }

  const PLANTS = RAW.map((row, index) => {
    const [idSap, denominazione, comune, indirizzo, lat, lng] = row;
    const id = plantId(idSap, denominazione, index);
    return applyCoordinatePair({
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
      stato: "DA FARE",
      statoGenerale: "DA FARE",
      done: false,
      tipo: "ORDINARIO",
      tipologiaIntervento: "ORDINARIO",
      attivitaLabel: "ORDINARIO",
      localFallback: true,
      sourceDataset: SOURCE_DATASET,
      sourceDatasetVersion: DATASET_VERSION
    }, { lat, lng });
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
    sourceDataset: SOURCE_DATASET,
    impiantiCount: PLANTS.length,
    totalPlants: PLANTS.length,
    impiantiFattiCount: 0,
    impiantiDaFareCount: PLANTS.length,
    localFallback: true
  };

  const collectionName = () => typeof getCommesseCollectionName === "function"
    ? getCommesseCollectionName()
    : "commesse";


  function installHistoricalCommesseResubscribe() {
    const GLOBAL = "HeraHistoricalCommesseResubscribe";
    const MAX_HISTORY_ATTEMPTS = 40;
    const HISTORY_RETRY_MS = 250;
    if (window[GLOBAL]?.installed) return window[GLOBAL];

    const state = {
      installed: true,
      attempts: 0,
      refreshed: false,
      lastReason: "",
      lastError: ""
    };
    let retryTimer = null;
    let authUnsubscribe = null;

    const currentUserReady = () => {
      try {
        return typeof currentUser !== "undefined" && Boolean(currentUser);
      } catch (_) {
        return false;
      }
    };

    const historicalRestoreReady = () => {
      const QueryPrototype = window.firebase?.firestore?.Query?.prototype;
      if (!QueryPrototype || typeof QueryPrototype.onSnapshot !== "function") return false;
      const restored = QueryPrototype.__heraActiveCommesseOriginalOnSnapshot;
      return typeof restored === "function" && QueryPrototype.onSnapshot === restored;
    };

    const schedule = (reason) => {
      if (state.refreshed || retryTimer) return;
      retryTimer = setTimeout(() => {
        retryTimer = null;
        refresh(reason);
      }, HISTORY_RETRY_MS);
    };

    const refresh = (reason = "runtime") => {
      if (state.refreshed) return true;
      state.attempts += 1;
      const ready = currentUserReady()
        && historicalRestoreReady()
        && typeof stopCommesseSubscription === "function"
        && typeof subscribeCommesse === "function";

      if (!ready) {
        if (state.attempts < MAX_HISTORY_ATTEMPTS) schedule(reason);
        return false;
      }

      try {
        stopCommesseSubscription();
        const request = subscribeCommesse();
        if (request && typeof request.catch === "function") {
          request.catch((error) => {
            state.lastError = String(error?.message || error || "");
            console.warn("[COMMESSE STORICHE] riavvio listener non riuscito", error);
          });
        }
        state.refreshed = true;
        state.lastReason = reason;
        if (authUnsubscribe) {
          try { authUnsubscribe(); } catch (_) {}
          authUnsubscribe = null;
        }
        return true;
      } catch (error) {
        state.lastError = String(error?.message || error || "");
        console.warn("[COMMESSE STORICHE] riavvio listener non riuscito", error);
        if (state.attempts < MAX_HISTORY_ATTEMPTS) schedule(reason);
        return false;
      }
    };

    window[GLOBAL] = {
      installed: true,
      refresh: () => refresh("manual"),
      getState: () => ({ ...state })
    };

    try {
      const authInstance = window.firebase?.auth?.();
      if (authInstance && typeof authInstance.onAuthStateChanged === "function") {
        authUnsubscribe = authInstance.onAuthStateChanged((user) => {
          if (user) schedule("auth");
        });
      }
    } catch (_) {}

    window.addEventListener("load", () => schedule("load"), { once: true });
    document.addEventListener("visibilitychange", () => {
      if (!document.hidden && !state.refreshed) schedule("visibility");
    });
    schedule("bootstrap");
    return window[GLOBAL];
  }

  function mergeObjectByPriority(target, source, priority) {
    Object.entries(source || {}).forEach(([field, value]) => {
      if (field === "sourceIds" || !hasValue(value)) return;
      if (priority > 0 || !hasValue(target[field])) target[field] = value;
    });
  }

  function clearMixedMarker(target) {
    target.isMixedOrdinaryExtraordinary = false;
    if (target.markerClass === "impianto-marker-mixed-yellow") delete target.markerClass;
    if (target.markerTone === "yellow") delete target.markerTone;
    if (String(target.markerColor || "").toLowerCase() === MIXED_MARKER_COLOR) delete target.markerColor;
    if (String(target.markerFillColor || "").toLowerCase() === MIXED_MARKER_COLOR) delete target.markerFillColor;
  }

  function applyKinds(target, kinds) {
    clearMixedMarker(target);
    const hasOrdinary = kinds.has("ORDINARIO");
    const hasExtraordinary = kinds.has("STRAORDINARIO");
    let label = "ORDINARIO";
    if (hasOrdinary && hasExtraordinary) label = "ORDINARIO E STRAORDINARIO";
    else if (hasExtraordinary) label = "STRAORDINARIO";

    target.tipo = label;
    target.tipologia = label;
    target.tipologiaIntervento = label;
    target.tipologiaLavorazione = label;
    target.attivitaLabel = label;

    if (hasOrdinary && hasExtraordinary) {
      target.isMixedOrdinaryExtraordinary = true;
      target.markerColor = MIXED_MARKER_COLOR;
      target.markerFillColor = MIXED_MARKER_COLOR;
      target.markerClass = "impianto-marker-mixed-yellow";
      target.markerTone = "yellow";
    }
  }

  function applyAggregateToPlant(target, aggregate) {
    if (!aggregate || aggregate.total <= 0) return target;
    target.numeroLavorazioni = aggregate.total;
    target.numeroLavorazioniFatte = aggregate.done;
    target.numeroLavorazioniDaFare = aggregate.pending;

    if (aggregate.allDone) {
      target.stato = "FATTO";
      target.statoGenerale = "FATTO";
      target.done = true;
    } else if (aggregate.partial && target.done !== true && !isDoneStatus(target.statoGenerale || target.stato)) {
      target.stato = "PARZIALMENTE FATTO";
      target.statoGenerale = "PARZIALMENTE FATTO";
      target.done = false;
    }
    return target;
  }

  function mergePlants(...collections) {
    const groups = [];
    const aliasToGroup = new Map();
    const uniqueAnagToGroup = new Map();

    function registerAnag(group, item) {
      const anag = anagAlias(item);
      if (!anag) return;
      if (!uniqueAnagToGroup.has(anag)) uniqueAnagToGroup.set(anag, group);
      else if (uniqueAnagToGroup.get(anag) !== group) uniqueAnagToGroup.set(anag, null);
    }

    function findGroup(item) {
      for (const alias of plantLookupAliases(item)) {
        const group = aliasToGroup.get(alias);
        if (group) return group;
      }
      if (!sapAlias(item) && extractPlantIds(item).length === 0) {
        const group = uniqueAnagToGroup.get(anagAlias(item));
        if (group) return group;
      }
      return null;
    }

    function registerAliases(group, item) {
      plantLookupAliases(item).forEach((alias) => aliasToGroup.set(alias, group));
      registerAnag(group, item);
      extractPlantIds(item).forEach((id) => group.sourceIds.add(id));
    }

    collections.forEach((collection, priority) => {
      const rows = Array.isArray(collection) ? collection : [collection];
      rows.filter(Boolean).forEach((item) => {
        let group = findGroup(item);
        if (!group) {
          group = {
            item: {},
            fallback: null,
            sourceIds: new Set(),
            kinds: new Set(),
            kindPriority: -1,
            workItems: []
          };
          groups.push(group);
        }

        const effectivePriority = item.localFallback === true ? 0 : priority;
        mergeObjectByPriority(group.item, item, effectivePriority);
        if (priority === 0 && !group.fallback) group.fallback = { ...item };
        registerAliases(group, item);

        const detected = detectKinds(item);
        if (detected.size > 0) {
          if (effectivePriority > group.kindPriority) {
            group.kinds = new Set(detected);
            group.kindPriority = effectivePriority;
          } else if (effectivePriority === group.kindPriority) {
            detected.forEach((kind) => group.kinds.add(kind));
          }
        }
      });
    });

    workItemsCache.forEach((workItem) => {
      let group = null;
      for (const alias of workItemPlantAliases(workItem)) {
        group = aliasToGroup.get(alias) || null;
        if (group) break;
      }
      if (!group) return;
      group.workItems.push(workItem);
      [workItem.physicalPlantId, workItem.impiantoId, baseMigrationSourceId(workItem.migrationSourceId)]
        .map(text)
        .filter(Boolean)
        .forEach((id) => group.sourceIds.add(id));
    });

    return groups.map((group) => {
      const out = {
        ...group.item,
        sourceIds: [...group.sourceIds]
      };
      const workKinds = new Set();
      group.workItems.forEach((item) => detectKinds(item).forEach((kind) => workKinds.add(kind)));
      applyKinds(out, workKinds.size ? workKinds : group.kinds);
      applyAggregateToPlant(out, deriveWorkAggregate(group.workItems));
      applyCoordinatePair(out, coordinatePairFrom(out) || coordinatePairFrom(group.fallback || {}));
      return out;
    }).sort((a, b) => {
      const left = Number(a.numeroProgressivo || a.numeroProgressivoImpianto) || 9999;
      const right = Number(b.numeroProgressivo || b.numeroProgressivoImpianto) || 9999;
      return left - right || identity(a).localeCompare(identity(b), "it");
    });
  }

  function signature(items) {
    return (Array.isArray(items) ? items : []).map((item) => JSON.stringify([
      identity(item),
      text(item.id),
      text(item.physicalPlantId),
      [...new Set(Array.isArray(item.sourceIds) ? item.sourceIds.map(text).filter(Boolean) : [])].sort(),
      text(item.idSap || item.idSAP),
      text(item.denominazione || item.nome),
      text(item.comune),
      text(item.indirizzo || item.via),
      coordinatePairFrom(item),
      text(item.attivitaLabel || item.tipologiaIntervento || item.tipo),
      text(item.statoGenerale || item.stato),
      item.done === true ? 1 : 0,
      item.localFallback === true ? 1 : 0,
      Number(item.numeroLavorazioni || 0),
      Number(item.numeroLavorazioniFatte || 0),
      Number(item.numeroLavorazioniDaFare || 0),
      text(item.markerClass),
      text(item.markerTone)
    ])).join(";");
  }

  function applyMergedVisiblePlants() {
    let cached = [];
    try {
      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.get) {
        cached = impiantiByCommessaId.get(COMMESSA_ID) || [];
      }
    } catch (_) {}

    let current = [];
    try {
      if (
        typeof selectedCommessaId !== "undefined"
        && selectedCommessaId === COMMESSA_ID
        && typeof currentImpianti !== "undefined"
        && Array.isArray(currentImpianti)
      ) {
        current = currentImpianti;
      }
    } catch (_) {}

    const merged = mergePlants(PLANTS, cached, current);
    try {
      if (
        typeof impiantiByCommessaId !== "undefined"
        && impiantiByCommessaId?.set
        && signature(cached) !== signature(merged)
      ) {
        impiantiByCommessaId.set(COMMESSA_ID, merged.map((item) => ({ ...item })));
      }
    } catch (_) {}

    try {
      if (
        typeof selectedCommessaId !== "undefined"
        && selectedCommessaId === COMMESSA_ID
        && typeof currentImpianti !== "undefined"
        && signature(current) !== signature(merged)
      ) {
        currentImpianti = merged.map((item) => ({ ...item }));
        if (typeof renderImpianti === "function") renderImpianti();
        if (typeof renderMap === "function") renderMap();
        if (typeof renderHeaderActivitySummary === "function") renderHeaderActivitySummary();
        if (typeof updateCommessaDashboard === "function") updateCommessaDashboard();
      }
    } catch (error) {
      console.warn("[INRETE Modena merge impianti UI]", error);
    }
    return merged;
  }

  async function refreshWorkKinds() {
    if (
      mixedRefreshRunning
      || typeof db === "undefined"
      || typeof auth === "undefined"
      || !auth.currentUser
      || (typeof document !== "undefined" && document.hidden)
    ) return false;

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

  function mergeCommessaLocally(existing) {
    const merged = { ...COMMESSA };
    Object.entries(existing || {}).forEach(([field, value]) => {
      if (hasValue(value) || value === false || value === 0) merged[field] = value;
    });
    return merged;
  }

  function ensureVisibleLocally() {
    try {
      let changed = false;
      if (typeof commesseById !== "undefined" && commesseById?.set) {
        const existing = commesseById.get(COMMESSA_ID);
        const mergedCommessa = mergeCommessaLocally(existing);
        if (!existing || JSON.stringify(existing) !== JSON.stringify(mergedCommessa)) {
          commesseById.set(COMMESSA_ID, mergedCommessa);
          changed = true;
        }
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

      if (
        typeof selectedCommessaId !== "undefined"
        && selectedCommessaId === COMMESSA_ID
        && typeof currentImpianti !== "undefined"
      ) {
        if (!Array.isArray(currentImpianti) || currentImpianti.length === 0) {
          currentImpianti = PLANTS.map((item) => ({ ...item }));
        }
        applyMergedVisiblePlants();
      }
    } catch (error) {
      console.warn("[INRETE Modena fallback UI]", error);
    }
  }

  function buildCanonicalIndex(existingCollections = []) {
    const aliasToPlant = new Map();
    const uniqueAnagToPlant = new Map();
    const register = (plant, item) => {
      plantLookupAliases(item).forEach((alias) => aliasToPlant.set(alias, plant));
      const anag = anagAlias(item);
      if (anag) {
        if (!uniqueAnagToPlant.has(anag)) uniqueAnagToPlant.set(anag, plant);
        else if (uniqueAnagToPlant.get(anag) !== plant) uniqueAnagToPlant.set(anag, null);
      }
    };

    PLANTS.forEach((plant) => register(plant, plant));
    existingCollections.flat().filter(Boolean).forEach((entry) => {
      const item = { id: entry.docId, ...(entry.data || {}) };
      let plant = null;
      for (const alias of plantLookupAliases(item)) {
        plant = aliasToPlant.get(alias) || null;
        if (plant) break;
      }
      if (!plant && !sapAlias(item) && extractPlantIds(item).length === 0) {
        plant = uniqueAnagToPlant.get(anagAlias(item)) || null;
      }
      if (plant) register(plant, item);
    });

    return {
      aliasToPlant,
      uniqueAnagToPlant,
      find(item = {}, options = {}) {
        const aliases = options.workItem ? workItemPlantAliases(item) : plantLookupAliases(item);
        for (const alias of aliases) {
          const plant = aliasToPlant.get(alias);
          if (plant) return plant;
        }
        if (!sapAlias(item) && (options.workItem || extractPlantIds(item).length === 0)) {
          return uniqueAnagToPlant.get(anagAlias(item)) || null;
        }
        return null;
      },
      register
    };
  }

  function groupWorkItemsByPlant(items, canonicalIndex) {
    const grouped = new Map();
    (Array.isArray(items) ? items : []).forEach((item) => {
      const plant = canonicalIndex.find(item, { workItem: true });
      if (!plant) return;
      if (!grouped.has(plant.id)) grouped.set(plant.id, []);
      grouped.get(plant.id).push(item);
    });
    return grouped;
  }

  function valuesEqual(left, right) {
    if (typeof left === "number" || typeof right === "number") {
      const a = Number(left);
      const b = Number(right);
      return Number.isFinite(a) && Number.isFinite(b) && Math.abs(a - b) < 0.0000001;
    }
    return left === right;
  }

  function buildExistingPlantPatch(current = {}, fallback = {}, aggregate = null, docId = "") {
    const patch = {};
    const setIfMissing = (field, value) => {
      if (!hasValue(current[field]) && hasValue(value)) patch[field] = value;
    };
    const setIfDifferent = (field, value) => {
      if (hasValue(value) && !valuesEqual(current[field], value)) patch[field] = value;
    };

    setIfMissing("id", docId || fallback.id);
    setIfMissing("commessaId", COMMESSA_ID);
    setIfMissing("physicalPlantId", docId || fallback.physicalPlantId || fallback.id);
    setIfMissing("migrationSourceId", docId || fallback.migrationSourceId || fallback.id);

    [
      "numeroProgressivo",
      "numeroProgressivoImpianto",
      "distretto",
      "idSap",
      "denominazione",
      "nome",
      "comune",
      "indirizzo",
      "via",
      "tipo",
      "tipologiaIntervento",
      "attivitaLabel"
    ].forEach((field) => setIfMissing(field, fallback[field]));

    const pair = coordinatePairFrom(current) || coordinatePairFrom(fallback);
    if (pair) {
      [
        ["gpsY", pair.lat],
        ["gpsX", pair.lng],
        ["latitudine", pair.lat],
        ["longitudine", pair.lng],
        ["latitude", pair.lat],
        ["longitude", pair.lng],
        ["lat", pair.lat],
        ["lng", pair.lng]
      ].forEach(([field, value]) => setIfDifferent(field, value));
    }

    setIfDifferent("localFallback", false);
    setIfDifferent("sourceDataset", SOURCE_DATASET);
    setIfDifferent("sourceDatasetVersion", DATASET_VERSION);

    if (aggregate && aggregate.total > 0) {
      setIfDifferent("numeroLavorazioni", aggregate.total);
      setIfDifferent("numeroLavorazioniFatte", aggregate.done);
      setIfDifferent("numeroLavorazioniDaFare", aggregate.pending);

      const currentDone = current.done === true || isDoneStatus(current.statoGenerale || current.stato);
      if (aggregate.allDone && !currentDone) {
        patch.stato = "FATTO";
        patch.statoGenerale = "FATTO";
        patch.done = true;
      } else if (aggregate.partial && !currentDone && !isPartialStatus(current.statoGenerale || current.stato)) {
        patch.stato = "PARZIALMENTE FATTO";
        patch.statoGenerale = "PARZIALMENTE FATTO";
        patch.done = false;
      }
    }

    return patch;
  }

  function buildNewPlantData(fallback, aggregate = null) {
    const data = { ...fallback, localFallback: false };
    applyCoordinatePair(data, coordinatePairFrom(fallback));
    applyAggregateToPlant(data, aggregate);
    return data;
  }

  function readSnapshotEntries(snapshot) {
    return snapshot.docs.map((doc) => ({
      docId: doc.id,
      data: doc.data() || {},
      ref: doc.ref
    }));
  }

  function finalPlantState(current, patch) {
    const combined = { ...current, ...patch };
    return {
      done: combined.done === true || isDoneStatus(combined.statoGenerale || combined.stato),
      state: combined
    };
  }

  async function createInFirestore(force = false) {
    if (running || typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser) return false;
    running = true;
    try {
      const ref = db.collection(collectionName()).doc(COMMESSA_ID);
      const existingParent = await ref.get();
      const existingParentData = existingParent.exists ? (existingParent.data() || {}) : {};
      if (
        !force
        && existingParent.exists
        && String(existingParentData.datasetVersion || "") === DATASET_VERSION
      ) return true;

      const [legacySnapshot, physicalSnapshot, workSnapshot] = await Promise.all([
        ref.collection("impianti").get(),
        ref.collection("impiantiFisici").get(),
        ref.collection("lavorazioni").get()
      ]);

      const legacyEntries = readSnapshotEntries(legacySnapshot);
      const physicalEntries = readSnapshotEntries(physicalSnapshot);
      workItemsCache = workSnapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));

      const canonicalIndex = buildCanonicalIndex([legacyEntries, physicalEntries]);
      const workItemsByPlant = groupWorkItemsByPlant(workItemsCache, canonicalIndex);
      const batch = db.batch();
      const serverTimestamp = firebase.firestore.FieldValue.serverTimestamp;
      const legacyFinalStates = new Map();

      function queueCollection(entries, collectionRef, collectFinalStates = false) {
        const matched = new Set();

        entries.forEach((entry) => {
          const item = { id: entry.docId, ...entry.data };
          const fallback = canonicalIndex.find(item);
          if (!fallback) return;
          matched.add(fallback.id);
          canonicalIndex.register(fallback, item);

          const aggregate = deriveWorkAggregate(workItemsByPlant.get(fallback.id) || []);
          const patch = buildExistingPlantPatch(entry.data, fallback, aggregate, entry.docId);
          if (Object.keys(patch).length > 0) {
            patch.updatedAt = serverTimestamp();
            batch.set(entry.ref, patch, { merge: true });
          }
          if (collectFinalStates) legacyFinalStates.set(fallback.id, finalPlantState(entry.data, patch));
        });

        PLANTS.forEach((fallback) => {
          if (matched.has(fallback.id)) return;
          const aggregate = deriveWorkAggregate(workItemsByPlant.get(fallback.id) || []);
          const data = buildNewPlantData(fallback, aggregate);
          data.createdAt = serverTimestamp();
          data.updatedAt = serverTimestamp();
          batch.set(collectionRef.doc(fallback.id), data, { merge: true });
          if (collectFinalStates) legacyFinalStates.set(fallback.id, finalPlantState({}, data));
        });
      }

      queueCollection(legacyEntries, ref.collection("impianti"), true);
      queueCollection(physicalEntries, ref.collection("impiantiFisici"), false);

      const doneCount = PLANTS.reduce((total, plant) => total + (legacyFinalStates.get(plant.id)?.done ? 1 : 0), 0);
      const parentPatch = {
        datasetVersion: DATASET_VERSION,
        sourceDataset: SOURCE_DATASET,
        localFallback: false,
        impiantiCount: PLANTS.length,
        totalPlants: PLANTS.length,
        impiantiFattiCount: doneCount,
        impiantiDaFareCount: Math.max(0, PLANTS.length - doneCount),
        operationalModelSyncedAt: serverTimestamp(),
        updatedAt: serverTimestamp()
      };

      [
        "id",
        "nome",
        "name",
        "codice",
        "code",
        "cliente",
        "categoria",
        "area",
        "stato",
        "attiva",
        "parentCommessaId",
        "excelModelVersion",
        "priceListVersion",
        "workItemsModelVersion",
        "percentualeRibassoGenerale",
        "nextImpiantoNumber"
      ].forEach((field) => {
        if (!hasValue(existingParentData[field]) && hasValue(COMMESSA[field])) parentPatch[field] = COMMESSA[field];
      });

      if (!existingParent.exists || !hasValue(existingParentData.createdAt)) parentPatch.createdAt = serverTimestamp();
      if (!hasValue(existingParentData.createdBy)) parentPatch.createdBy = auth.currentUser.uid;
      if (!hasValue(existingParentData.creatoDa)) parentPatch.creatoDa = auth.currentUser.email || "";

      batch.set(ref, parentPatch, { merge: true });
      await batch.commit();
      applyMergedVisiblePlants();
      console.info("[INRETE Modena] sincronizzazione non distruttiva completata", {
        count: PLANTS.length,
        doneCount,
        datasetVersion: DATASET_VERSION
      });
      return true;
    } finally {
      running = false;
    }
  }

  async function tryRun() {
    ensureVisibleLocally();
    attempts += 1;
    try {
      if (typeof auth === "undefined" || typeof db === "undefined" || !auth.currentUser) {
        throw new Error("Firebase non pronto");
      }
      if (await createInFirestore(false)) {
        await refreshWorkKinds();
        return;
      }
    } catch (error) {
      console.warn(
        "[INRETE Modena] sincronizzazione Firestore non riuscita, resta attivo il fallback UI",
        attempts,
        error?.message || error
      );
    }
    if (attempts < MAX_ATTEMPTS) setTimeout(tryRun, RETRY_MS);
  }

  function setWorkItemsForTesting(items) {
    workItemsCache = Array.isArray(items) ? items.map((item) => ({ ...item })) : [];
  }

  window.createInreteModenaAugust2026 = (options = {}) => createInFirestore(options.force === true);
  window.refreshInreteModenaMixedWork = refreshWorkKinds;
  window.INRETE_MODENA_AUGUST_2026 = {
    commessa: COMMESSA,
    plants: PLANTS,
    ensureVisibleLocally,
    refreshWorkKinds,
    mergePlants,
    testing: {
      detectKinds: (item) => [...detectKinds(item)],
      deriveWorkAggregate,
      coordinatePairFrom,
      signature,
      buildExistingPlantPatch,
      setWorkItems: setWorkItemsForTesting
    }
  };

  installHistoricalCommesseResubscribe();

  setInterval(ensureVisibleLocally, 2000);
  setInterval(() => {
    try {
      if (
        typeof selectedCommessaId !== "undefined"
        && selectedCommessaId === COMMESSA_ID
        && (typeof document === "undefined" || !document.hidden)
      ) refreshWorkKinds();
    } catch (_) {}
  }, MIXED_REFRESH_MS);

  if (typeof auth !== "undefined" && auth?.onAuthStateChanged) {
    auth.onAuthStateChanged((user) => {
      if (user) setTimeout(tryRun, 250);
    });
  }

  window.addEventListener("load", () => {
    ensureVisibleLocally();
    setTimeout(tryRun, 500);
    setTimeout(refreshWorkKinds, 1200);
  });

  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      ensureVisibleLocally();
      try {
        if (typeof selectedCommessaId !== "undefined" && selectedCommessaId === COMMESSA_ID) refreshWorkKinds();
      } catch (_) {}
    }
  });

  setTimeout(ensureVisibleLocally, 100);
  setTimeout(tryRun, 1000);
})();

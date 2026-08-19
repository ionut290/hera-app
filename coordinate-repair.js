(function exposeCoordinateRepair(root) {
  "use strict";

  const ITALY_BOUNDS = Object.freeze({
    minLat: 35,
    maxLat: 48.8,
    minLng: 5,
    maxLng: 20
  });

  const LATITUDE_KEYS = Object.freeze([
    "coordinateLatitudineOriginale", "latitudine", "gpsY", "latitude", "lat",
    "coordinateY", "y", "Coordinate GPS Y / Latitudine", "Coordinate GPS Y",
    "GPS Y", "Coordinate GPS(Y)", "GPS(Y)", "Coordinata Y", "Latitudine GPS",
    "coordinate_gps_y", "gps_y", "coordY", "coord_y"
  ]);
  const LONGITUDE_KEYS = Object.freeze([
    "coordinateLongitudineOriginale", "longitudine", "gpsX", "longitude", "lng", "lon",
    "coordinateX", "x", "Coordinate GPS X / Longitudine", "Coordinate GPS X",
    "GPS X", "Coordinate GPS(X)", "GPS(X)", "Coordinata X", "Longitudine GPS",
    "coordinate_gps_x", "gps_x", "coordX", "coord_x"
  ]);
  const PAIR_KEYS = Object.freeze([
    "coordinate", "coordinates", "gps", "coordinateGps", "coordinateGPS", "Coordinate GPS",
    "coordinateUnica", "coordinateOriginale", "Coordinate GPS(X)/GPS(Y)",
    "Coordinate GPS X/Y", "latLng", "latlng", "position", "posizione"
  ]);
  const NESTED_KEYS = Object.freeze([
    "coordinate", "coordinates", "gps", "coordinateGps", "coordinateGPS", "geoPoint",
    "geopoint", "location", "posizione", "position", "coords"
  ]);
  const EXTRA_FIELD_KEYS = Object.freeze([
    "extraFields", "campiExtra", "additionalFields", "rawFields"
  ]);

  function cleanRaw(value) {
    return value == null ? "" : String(value).trim();
  }

  function parseSingle(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    const raw = cleanRaw(value);
    if (!raw) return null;
    let normalized = raw
      .replace(/[°º'"`]/g, "")
      .replace(/\s+/g, "");
    if (normalized.includes(",") && normalized.includes(".")) {
      if (normalized.lastIndexOf(",") > normalized.lastIndexOf(".")) {
        normalized = normalized.replace(/\./g, "").replace(/,/g, ".");
      } else {
        normalized = normalized.replace(/,/g, "");
      }
    } else {
      normalized = normalized.replace(/,/g, ".");
    }
    if (!/^[-+]?\d+(?:\.\d+)?$/.test(normalized)) return null;
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function extractPair(value) {
    const raw = cleanRaw(value);
    if (!raw) return null;
    const matches = raw.match(/[-+]?\d{1,3}(?:[.,]\d+)?/g) || [];
    if (matches.length !== 2) return null;
    const first = parseSingle(matches[0]);
    const second = parseSingle(matches[1]);
    return first == null || second == null ? null : [first, second];
  }

  function isValidLatitude(value) {
    return Number.isFinite(value) && value >= -90 && value <= 90 && value !== 0;
  }

  function isValidLongitude(value) {
    return Number.isFinite(value) && value >= -180 && value <= 180 && value !== 0;
  }

  function looksItalian(lat, lng) {
    return lat >= ITALY_BOUNDS.minLat
      && lat <= ITALY_BOUNDS.maxLat
      && lng >= ITALY_BOUNDS.minLng
      && lng <= ITALY_BOUNDS.maxLng;
  }

  function diagnose(latitudeRaw, longitudeRaw) {
    const rawLatitude = cleanRaw(latitudeRaw);
    const rawLongitude = cleanRaw(longitudeRaw);
    let latitude = parseSingle(rawLatitude);
    let longitude = parseSingle(rawLongitude);
    let repairType = "";

    if (latitude == null || longitude == null) {
      const pair = extractPair(rawLatitude) || extractPair(rawLongitude);
      if (pair) {
        latitude = pair[0];
        longitude = pair[1];
        repairType = "PAIR";
      }
    }

    const directValid = isValidLatitude(latitude) && isValidLongitude(longitude);
    const swappedValid = isValidLatitude(longitude) && isValidLongitude(latitude);
    const clearlySwapped = swappedValid && (
      !directValid
      || (!looksItalian(latitude, longitude) && looksItalian(longitude, latitude))
    );
    if (clearlySwapped) {
      [latitude, longitude] = [longitude, latitude];
      repairType = "SWAPPED";
    }

    const valid = isValidLatitude(latitude) && isValidLongitude(longitude);
    if (!valid) {
      const missing = !rawLatitude && !rawLongitude;
      return {
        valid: false,
        repaired: false,
        latitude: null,
        longitude: null,
        rawLatitude,
        rawLongitude,
        status: missing ? "MISSING" : "INVALID",
        message: missing
          ? "Coordinate GPS mancanti"
          : `Coordinate GPS non valide: ${rawLatitude || "—"} / ${rawLongitude || "—"}`
      };
    }

    if (repairType) {
      return {
        valid: true,
        repaired: true,
        latitude,
        longitude,
        rawLatitude,
        rawLongitude,
        status: "REPAIRED",
        message: repairType === "SWAPPED"
          ? "Coordinate GPS invertite e corrette automaticamente"
          : "Coordinate GPS unite e separate automaticamente"
      };
    }

    return {
      valid: true,
      repaired: false,
      latitude,
      longitude,
      rawLatitude,
      rawLongitude,
      status: "VALID",
      message: ""
    };
  }

  function normalizeKey(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]/g, "");
  }

  function present(value) {
    return value !== undefined && value !== null && String(value).trim() !== "";
  }

  function valuesForKeys(record, keys, normalizedNames) {
    if (!record || typeof record !== "object") return [];
    const values = [];
    keys.forEach((key) => {
      if (present(record[key])) values.push({ value: record[key], key });
    });
    Object.entries(record).forEach(([key, value]) => {
      if (!present(value)) return;
      const normalized = normalizeKey(key);
      if (normalizedNames.has(normalized) && !values.some((entry) => entry.key === key)) {
        values.push({ value, key });
      }
    });
    return values;
  }

  const NORMALIZED_LATITUDE_KEYS = new Set(LATITUDE_KEYS.map(normalizeKey));
  const NORMALIZED_LONGITUDE_KEYS = new Set(LONGITUDE_KEYS.map(normalizeKey));
  const NORMALIZED_PAIR_KEYS = new Set(PAIR_KEYS.map(normalizeKey));

  function resolveRecord(record, depth = 0) {
    if (!record || typeof record !== "object") {
      return { ...diagnose("", ""), source: "" };
    }

    const latitudeValues = valuesForKeys(record, LATITUDE_KEYS, NORMALIZED_LATITUDE_KEYS);
    const longitudeValues = valuesForKeys(record, LONGITUDE_KEYS, NORMALIZED_LONGITUDE_KEYS);
    let firstInvalid = null;

    for (const latitudeEntry of latitudeValues) {
      for (const longitudeEntry of longitudeValues) {
        const result = diagnose(latitudeEntry.value, longitudeEntry.value);
        if (result.valid) {
          return { ...result, source: `${latitudeEntry.key} + ${longitudeEntry.key}` };
        }
        if (!firstInvalid && result.status === "INVALID") firstInvalid = result;
      }
    }

    // Alcuni archivi INRETE salvano la coppia X/Y completa in uno solo dei
    // campi storici di latitudine o longitudine. Il parser `diagnose` sa già
    // separare e invertire la coppia, ma prima questa strada non veniva
    // tentata da `resolveRecord` quando il campo gemello era vuoto.
    for (const latitudeEntry of latitudeValues) {
      const result = diagnose(latitudeEntry.value, "");
      if (result.valid) {
        return { ...result, source: `${latitudeEntry.key} (coordinate unite)` };
      }
      if (!firstInvalid && result.status === "INVALID") firstInvalid = result;
    }
    for (const longitudeEntry of longitudeValues) {
      const result = diagnose("", longitudeEntry.value);
      if (result.valid) {
        return { ...result, source: `${longitudeEntry.key} (coordinate unite)` };
      }
      if (!firstInvalid && result.status === "INVALID") firstInvalid = result;
    }

    for (const key of NESTED_KEYS) {
      const nested = record[key];
      if (!nested || typeof nested !== "object") continue;
      const nestedLat = nested.latitude ?? nested.lat ?? nested.gpsY ?? nested.y;
      const nestedLng = nested.longitude ?? nested.lng ?? nested.lon ?? nested.gpsX ?? nested.x;
      const result = diagnose(nestedLat, nestedLng);
      if (result.valid) return { ...result, source: `${key} (oggetto)` };
      if (!firstInvalid && result.status === "INVALID") firstInvalid = result;
    }

    const pairValues = valuesForKeys(record, PAIR_KEYS, NORMALIZED_PAIR_KEYS);
    for (const pairEntry of pairValues) {
      const result = diagnose(pairEntry.value, "");
      if (result.valid) return { ...result, source: pairEntry.key };
      if (!firstInvalid && result.status === "INVALID") firstInvalid = result;
    }

    // Le importazioni storiche salvavano talvolta le colonne originali del
    // foglio (per esempio coordinategpsy/coordinategpsx) dentro extraFields.
    // Limitiamo la ricorsione ai contenitori noti per non scandire dati non
    // pertinenti e per mantenere invariato il percorso di lettura normale.
    if (depth < 2) {
      for (const key of EXTRA_FIELD_KEYS) {
        const nested = record[key];
        if (!nested || typeof nested !== "object" || Array.isArray(nested)) continue;
        const result = resolveRecord(nested, depth + 1);
        if (result.valid) return { ...result, source: `${key}.${result.source || "coordinate"}` };
        if (!firstInvalid && result.status === "INVALID") firstInvalid = result;
      }
    }

    return { ...(firstInvalid || diagnose("", "")), source: "" };
  }

  function normalizeRecord(record) {
    if (!record || typeof record !== "object") return record;
    const result = resolveRecord(record);
    if (!result.valid) return { ...record };

    const normalized = {
      ...record,
      gpsY: result.latitude,
      gpsX: result.longitude,
      latitudine: result.latitude,
      longitudine: result.longitude,
      latitude: result.latitude,
      longitude: result.longitude,
      coordinateStatus: result.status,
      coordinateIssue: result.message || "",
      coordinateResolvedFrom: result.source || record.coordinateResolvedFrom || ""
    };
    if (!present(normalized.coordinateLatitudineOriginale)) {
      normalized.coordinateLatitudineOriginale = result.rawLatitude || result.latitude;
    }
    if (!present(normalized.coordinateLongitudineOriginale)) {
      normalized.coordinateLongitudineOriginale = result.rawLongitude || result.longitude;
    }
    return normalized;
  }

  root.HeraCoordinateRepair = Object.freeze({
    diagnose,
    parseSingle,
    extractPair,
    isValidLatitude,
    isValidLongitude,
    resolveRecord,
    normalizeRecord,
    latitudeKeys: LATITUDE_KEYS,
    longitudeKeys: LONGITUDE_KEYS
  });
})(window);

(function installCoordinateReadRuntime(root) {
  "use strict";

  if (root.HeraCoordinateReadRuntime?.installed) return;
  const tools = root.HeraCoordinateRepair;
  if (!tools?.normalizeRecord || !tools?.resolveRecord) return;

  const state = {
    normalizedRows: 0,
    repairedRows: 0,
    physicalReads: 0,
    physicalMatches: 0,
    physicalRerenders: 0,
    lastCommessaId: "",
    errors: []
  };
  const physicalCache = new Map();
  const physicalRequests = new Map();
  const rerenderSignatures = new Map();

  function normalizeRows(rows) {
    if (!Array.isArray(rows)) return rows;
    return rows.map((row) => {
      const before = tools.resolveRecord(row);
      const normalized = tools.normalizeRecord(row);
      const after = tools.resolveRecord(normalized);
      if (after.valid) state.normalizedRows += 1;
      if (!before.valid && after.valid) state.repairedRows += 1;
      return normalized;
    });
  }

  function getSelectedCommessaId() {
    try {
      return String(typeof selectedCommessaId !== "undefined" ? selectedCommessaId : root.selectedCommessaId || "").trim();
    } catch (_) {
      return String(root.selectedCommessaId || "").trim();
    }
  }

  function collectionName() {
    try {
      return String(typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse").trim() || "commesse";
    } catch (_) {
      return "commesse";
    }
  }

  function commessaRecord(commessaId) {
    try {
      if (typeof commesseById !== "undefined" && commesseById?.get) return commesseById.get(commessaId) || null;
    } catch (_) {}
    try {
      if (root.commesseById?.get) return root.commesseById.get(commessaId) || null;
    } catch (_) {}
    return null;
  }

  function isInreteCommessa(commessaId, rows) {
    const commessa = commessaRecord(commessaId);
    try {
      if (root.InreteWorkItemsV2?.isInreteCommessa?.(commessa)) return true;
    } catch (_) {}
    const text = [
      commessa?.nome, commessa?.codice, commessa?.categoria, commessa?.tipo,
      document.getElementById("commessa-focus-label")?.textContent,
      ...(Array.isArray(rows) ? rows.slice(0, 5).flatMap((row) => [row?.commessaNome, row?.commessa, row?.categoria]) : [])
    ].filter(Boolean).join(" ").toUpperCase();
    return text.includes("INRETE");
  }

  function normalizedText(value) {
    return String(value || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function firstValue(record, keys) {
    for (const key of keys) {
      const value = record?.[key];
      if (value !== undefined && value !== null && String(value).trim()) return String(value).trim();
    }
    return "";
  }

  function identityKeys(record) {
    const keys = new Set();
    [
      record?.id,
      record?.physicalPlantId,
      record?.impiantoFisicoId,
      record?.sourceImpiantoId,
      record?.plantId,
      record?.migrationSourceId,
      record?.impiantoId
    ].filter(Boolean).forEach((value) => {
      const raw = String(value).trim();
      if (!raw) return;
      keys.add(`id:${raw}`);
      ["::", "__"].forEach((separator) => {
        const separatorIndex = raw.indexOf(separator);
        if (separatorIndex > 0) keys.add(`id:${raw.slice(0, separatorIndex)}`);
      });
    });
    const sap = firstValue(record, ["idSap", "idSAP", "ID SAP", "idsap", "sap"]);
    if (sap) keys.add(`sap:${normalizedText(sap)}`);
    const name = firstValue(record, ["denominazione", "denominazioneImpianto", "Denominazione Impianto", "nome", "impianto"]);
    const comune = firstValue(record, ["comune", "Comune", "ubicazione", "localita", "località"]);
    const address = firstValue(record, ["indirizzo", "descrizioneVia", "Via e civico di ubicazione Impianto", "via", "address"]);
    if (name) {
      const normalizedName = normalizedText(name);
      const normalizedComune = normalizedText(comune);
      const normalizedAddress = normalizedText(address);
      if (normalizedComune || normalizedAddress) {
        keys.add(`name-address:${normalizedName}|${normalizedComune}|${normalizedAddress}`);
      }
      if (normalizedComune) keys.add(`name-comune:${normalizedName}|${normalizedComune}`);
      keys.add(`name:${normalizedName}`);
    }
    return keys;
  }

  function physicalIndex(rows) {
    const index = new Map();
    rows.forEach((row) => identityKeys(row).forEach((key) => {
      if (!index.has(key)) {
        index.set(key, row);
      } else if (index.get(key) !== row) {
        // Un nome/SAP duplicato non deve mai collegare coordinate al sito
        // sbagliato. I fallback meno specifici sono usati solo se univoci.
        index.set(key, null);
      }
    }));
    return index;
  }

  function matchPhysical(row, index) {
    for (const key of identityKeys(row)) {
      const match = index.get(key);
      if (match) return match;
    }
    return null;
  }

  function firestoreDb() {
    try {
      if (typeof db !== "undefined" && db?.collection) return db;
    } catch (_) {}
    return root.db?.collection ? root.db : null;
  }

  async function loadPhysicalRows(commessaId) {
    if (physicalCache.has(commessaId)) return physicalCache.get(commessaId);
    if (physicalRequests.has(commessaId)) return physicalRequests.get(commessaId);
    const promise = (async () => {
      const firestore = firestoreDb();
      if (!firestore || !commessaId) return [];
      const snapshot = await firestore.collection(collectionName()).doc(commessaId).collection("impiantiFisici").get();
      state.physicalReads += 1;
      const rows = snapshot.docs.map((doc) => tools.normalizeRecord({ id: doc.id, ...doc.data() }));
      physicalCache.set(commessaId, rows);
      return rows;
    })().catch((error) => {
      state.errors.push(String(error?.message || error));
      console.warn("Coordinate impianti fisici non caricate:", error);
      return [];
    }).finally(() => physicalRequests.delete(commessaId));
    physicalRequests.set(commessaId, promise);
    return promise;
  }

  function updateMemory(commessaId, rows) {
    try {
      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.set) {
        impiantiByCommessaId.set(commessaId, rows);
      }
    } catch (_) {}
    try {
      const selected = getSelectedCommessaId();
      if (selected === commessaId && typeof currentImpianti !== "undefined") currentImpianti = rows;
    } catch (_) {}
    try {
      if (getSelectedCommessaId() === commessaId) root.currentImpianti = rows;
    } catch (_) {}
  }

  async function enrichMissingFromPhysical(commessaId, rows, previousDoneSignatureRef, originalRender) {
    if (!commessaId || !Array.isArray(rows) || !rows.length || !isInreteCommessa(commessaId, rows)) return;
    const missing = rows.filter((row) => !tools.resolveRecord(row).valid);
    if (!missing.length) return;

    const physical = await loadPhysicalRows(commessaId);
    if (!physical.length || getSelectedCommessaId() !== commessaId) return;
    const index = physicalIndex(physical);
    let matches = 0;
    const enriched = rows.map((row) => {
      if (tools.resolveRecord(row).valid) return row;
      const physicalRow = matchPhysical(row, index);
      const physicalCoordinates = tools.resolveRecord(physicalRow);
      if (!physicalRow || !physicalCoordinates.valid) return row;
      matches += 1;
      return tools.normalizeRecord({
        ...row,
        gpsY: physicalCoordinates.latitude,
        gpsX: physicalCoordinates.longitude,
        latitudine: physicalCoordinates.latitude,
        longitudine: physicalCoordinates.longitude,
        coordinateResolvedFrom: `impiantiFisici:${physicalRow.id || "collegamento"}`
      });
    });
    if (!matches) return;

    const signature = enriched.map((row) => {
      const coordinate = tools.resolveRecord(row);
      return `${row?.id || row?.physicalPlantId || ""}:${coordinate.latitude || ""}:${coordinate.longitude || ""}`;
    }).join("|");
    if (rerenderSignatures.get(commessaId) === signature) return;
    rerenderSignatures.set(commessaId, signature);
    state.physicalMatches += matches;
    state.physicalRerenders += 1;
    state.lastCommessaId = commessaId;
    updateMemory(commessaId, enriched);
    originalRender.call(root, enriched, previousDoneSignatureRef);
  }

  function patchRowsFunction(name, original, assignLexical) {
    if (typeof original !== "function" || original.__heraCoordinateReadPatched) return;
    const patched = function patchedCoordinateRows() {
      const result = original.apply(this, arguments);
      return Array.isArray(result) ? normalizeRows(result) : result;
    };
    patched.__heraCoordinateReadPatched = true;
    patched.__heraCoordinateOriginal = original;
    root[name] = patched;
    try { assignLexical?.(patched); } catch (_) {}
  }

  let cachedReader = root.getCommessaCachedImpianti;
  try {
    if (typeof getCommessaCachedImpianti === "function") cachedReader = getCommessaCachedImpianti;
  } catch (_) {}
  patchRowsFunction("getCommessaCachedImpianti", cachedReader, (patched) => { getCommessaCachedImpianti = patched; });

  let combiner = root.combineImpiantiForView;
  try {
    if (typeof combineImpiantiForView === "function") combiner = combineImpiantiForView;
  } catch (_) {}
  patchRowsFunction("combineImpiantiForView", combiner, (patched) => { combineImpiantiForView = patched; });

  let originalRender = root.renderImpiantiAfterRemoteSync;
  try {
    if (typeof renderImpiantiAfterRemoteSync === "function") originalRender = renderImpiantiAfterRemoteSync;
  } catch (_) {}
  if (typeof originalRender === "function" && !originalRender.__heraCoordinateReadPatched) {
    const patchedRender = function renderImpiantiAfterCoordinateRepair(rawImpianti, previousDoneSignatureRef) {
      const normalized = normalizeRows(rawImpianti);
      const result = originalRender.call(this, normalized, previousDoneSignatureRef);
      const commessaId = getSelectedCommessaId();
      void enrichMissingFromPhysical(commessaId, normalized, previousDoneSignatureRef, originalRender);
      return result;
    };
    patchedRender.__heraCoordinateReadPatched = true;
    patchedRender.__heraCoordinateOriginal = originalRender;
    root.renderImpiantiAfterRemoteSync = patchedRender;
    try { renderImpiantiAfterRemoteSync = patchedRender; } catch (_) {}
  }

  root.HeraCoordinateReadRuntime = Object.freeze({
    installed: true,
    version: "2.1.0",
    normalizeRows,
    resolveRecord: tools.resolveRecord,
    matchPhysicalRecord: (row, physicalRows) => matchPhysical(row, physicalIndex(Array.isArray(physicalRows) ? physicalRows : [])),
    getState: () => ({
      ...state,
      cachedPhysicalCommesse: physicalCache.size,
      pendingPhysicalReads: physicalRequests.size
    }),
    clearPhysicalCache: (commessaId = "") => {
      if (commessaId) physicalCache.delete(String(commessaId));
      else physicalCache.clear();
    }
  });
})(window);

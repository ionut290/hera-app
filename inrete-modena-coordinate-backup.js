(function exposeInreteModenaCoordinateBackup(root) {
  "use strict";

  const SOURCE = "CONTABILITA_INRETE_MODENA_2026-08-12.xlsx";
  const BACKUP_ROWS = Object.freeze([
    { sap: "3430707", name: "REMI CASTELFRANCO", lat: 44.58931, lng: 11.04352 },
    { sap: "3430051", name: "REMI S.CESARIO", lat: 44.56565, lng: 11.03273 },
    { sap: "3472968", name: "REMI SAVIGNANO", lat: 44.48468, lng: 11.03210 },
    { sap: "3471231", name: "REMI SPEZZANO", lat: 44.53632, lng: 10.83482 },
    { sap: "3470478", name: "REMI FORMIGINE", lat: 44.56882, lng: 10.84307 },
    { sap: "3475570", name: "REMI MAGRETA", lat: 44.59869, lng: 10.79012 },
    { sap: "3471127", name: "REMI CASINALBO", lat: 44.59549, lng: 10.84845 },
    { sap: "3471152", name: "REMI MARANELLO", lat: 44.52638, lng: 10.85457 },
    { sap: "3470982", name: "REMI POZZA", lat: 44.53058, lng: 10.89367 },
    { sap: "3471269", name: "REMI BRAIDA", lat: 44.54703, lng: 10.79935 },
    { sap: "area adiacente alla REMI", name: "STOCCAGGIO TUBI BRAIDA", lat: 44.54703, lng: 10.79935 },
    { sap: "3470559", name: "REMI INDIPENDENZA", lat: 44.53781, lng: 10.77355 },
    { sap: "3471102", name: "REMI PONTE FOSSA", lat: 44.56558, lng: 10.80002 },
    { sap: "3471279", name: "REMI UBERSETTO", lat: 44.55092, lng: 10.86203 },
    { sap: "3429985", name: "REMI CASTELNUOVO", lat: 44.53866, lng: 10.94932 },
    { sap: "3430270", name: "REMI CA' DI SOLA", lat: 44.53481, lng: 10.95387 },
    { sap: "3430320", name: "REMI SOLIGNANO", lat: 44.53274, lng: 10.91688 },
    { sap: "3425784", name: "REMI S.CLEMENTE", lat: 44.70429, lng: 10.99211 },
    { sap: "3426714", name: "REMI S. DAMASO", lat: 44.61267, lng: 10.97673 },
    { sap: "3430248", name: "REMI SUD", lat: 44.60603, lng: 10.91114 },
    { sap: "3430604", name: "REMI CASTELLARO", lat: 44.54071, lng: 11.02894 },
    { sap: "3430547", name: "REMI S.VITO", lat: 44.53941, lng: 11.01387 },
    { sap: "3430221", name: "REMI VIGNOLA", lat: 44.48418, lng: 11.02692 },
    { sap: "3476117", name: "REMI BERZIGALA", lat: 44.39255, lng: 10.80896 },
    { sap: "3473527", name: "REMI PAVULLO CASTELLO", lat: 44.32698, lng: 10.83440 },
    { sap: "3473528", name: "REMI S.ANTONIO", lat: 44.36767, lng: 10.83285 },
    { sap: "3471537", name: "REMI S. DALMAZIO", lat: 44.42214, lng: 10.85448 },
    { sap: "3475075", name: "REMI MONTECENERE", lat: 44.31195, lng: 10.78122 },
    { sap: "3426336", name: "GRMI UMF07 _ EXPORT CERAM", lat: 44.37793, lng: 10.61998 },
    { sap: "3426335", name: "GRMI UMF08 _ EXPORT CERAM", lat: 44.37797, lng: 10.62001 }
  ]);

  const persisted = new Set();
  let persistenceRunning = null;

  function normalize(value) {
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
      if (value !== undefined && value !== null && String(value).trim()) return value;
    }
    return "";
  }

  function validPair(lat, lng) {
    const latitude = Number(lat);
    const longitude = Number(lng);
    return Number.isFinite(latitude) && latitude >= -90 && latitude <= 90 && latitude !== 0
      && Number.isFinite(longitude) && longitude >= -180 && longitude <= 180 && longitude !== 0;
  }

  function readCoordinates(record) {
    const match = [
      [record?.gpsY, record?.gpsX],
      [record?.latitudine, record?.longitudine],
      [record?.latitude ?? record?.lat, record?.longitude ?? record?.lng ?? record?.lon]
    ].find(([lat, lng]) => validPair(lat, lng));
    return match ? { lat: Number(match[0]), lng: Number(match[1]) } : null;
  }

  function hasCoordinates(record) {
    return Boolean(readCoordinates(record));
  }

  const bySap = new Map();
  const byName = new Map();
  BACKUP_ROWS.forEach((row) => {
    const sap = normalize(row.sap);
    const name = normalize(row.name);
    if (sap) bySap.set(sap, row);
    if (name) byName.set(name, row);
  });

  function findBackupRow(record) {
    const sap = normalize(firstValue(record, ["idSap", "idSAP", "ID SAP", "codiceHera", "sap"]));
    if (sap && bySap.has(sap)) return bySap.get(sap);
    const name = normalize(firstValue(record, ["denominazione", "denominazioneImpianto", "nome", "impianto"]));
    return name ? byName.get(name) || null : null;
  }

  function coordinatePayload(backup) {
    const payload = {
      gpsY: backup.lat,
      gpsX: backup.lng,
      latitudine: backup.lat,
      longitudine: backup.lng,
      latitude: backup.lat,
      longitude: backup.lng,
      coordinateStatus: "RESTORED_FROM_BACKUP",
      coordinateIssue: "",
      coordinateResolvedFrom: SOURCE,
      coordinateLatitudineOriginale: backup.lat,
      coordinateLongitudineOriginale: backup.lng
    };
    try {
      if (root.firebase?.firestore?.FieldValue?.serverTimestamp) {
        payload.coordinateRestoredAt = root.firebase.firestore.FieldValue.serverTimestamp();
      }
    } catch (_) {}
    try {
      const user = typeof currentUser !== "undefined" ? currentUser : root.currentUser;
      payload.coordinateRestoredBy = user?.email || user?.uid || "";
    } catch (_) {}
    return payload;
  }

  function restoreRecord(record) {
    if (!record || typeof record !== "object") return false;
    const existing = readCoordinates(record);
    if (existing) {
      if (!validPair(record.gpsY, record.gpsX)) {
        record.gpsY = existing.lat;
        record.gpsX = existing.lng;
      }
      return false;
    }
    const backup = findBackupRow(record);
    if (!backup) return false;
    Object.assign(record, coordinatePayload(backup));
    return true;
  }

  function restoreRows(rows) {
    if (!Array.isArray(rows)) return rows;
    rows.forEach(restoreRecord);
    return rows;
  }

  function getRuntimeDb() {
    try { if (typeof db !== "undefined" && db) return db; } catch (_) {}
    return root.db || null;
  }

  function getSelectedCommessaId() {
    try { if (typeof selectedCommessaId !== "undefined" && selectedCommessaId) return String(selectedCommessaId); } catch (_) {}
    return String(root.selectedCommessaId || "");
  }

  function getCurrentRows() {
    try { if (typeof currentImpianti !== "undefined" && Array.isArray(currentImpianti)) return currentImpianti; } catch (_) {}
    return Array.isArray(root.currentImpianti) ? root.currentImpianti : [];
  }

  function isModenaContext(rows, commessaId) {
    let commessa = null;
    try {
      const map = typeof commesseById !== "undefined" ? commesseById : root.commesseById;
      if (map?.get) commessa = map.get(commessaId) || null;
    } catch (_) {}
    const text = normalize([
      commessa?.nome, commessa?.codice, commessa?.categoria, commessa?.tipo,
      commessa?.commessaPadre, commessa?.parentName
    ].filter(Boolean).join(" "));
    if (text.includes("inrete") && text.includes("modena")) return true;
    const matchingRows = (rows || []).filter((row) => findBackupRow(row)).length;
    return matchingRows >= 5;
  }

  async function updateExistingRef(ref, backup, key) {
    if (persisted.has(key)) return "cached";
    const snap = await ref.get();
    if (!snap.exists) return "missing";
    const current = snap.data() || {};
    if (hasCoordinates(current)) {
      persisted.add(key);
      return "already-valid";
    }
    await ref.set(coordinatePayload(backup), { merge: true });
    persisted.add(key);
    return "updated";
  }

  async function persistRestoredRows(rows) {
    const runtimeDb = getRuntimeDb();
    const commessaId = getSelectedCommessaId();
    if (!runtimeDb || !commessaId || !Array.isArray(rows) || !rows.length) return { updated: 0 };
    if (!isModenaContext(rows, commessaId)) return { updated: 0 };
    if (persistenceRunning) return persistenceRunning;

    persistenceRunning = (async () => {
      let updated = 0;
      const base = runtimeDb.collection("globalCommesse").doc(commessaId);
      for (const record of rows) {
        const id = String(record?.id || record?.impiantoId || "").trim();
        const backup = findBackupRow(record);
        if (!id || !backup) continue;
        const targets = [
          ["impianti", base.collection("impianti").doc(id)],
          ["impiantiFisici", base.collection("impiantiFisici").doc(id)]
        ];
        for (const [collectionName, ref] of targets) {
          const key = `${commessaId}/${collectionName}/${id}`;
          try {
            const result = await updateExistingRef(ref, backup, key);
            if (result === "updated") updated += 1;
          } catch (error) {
            console.warn(`Ripristino coordinate INRETE Modena non salvato su ${key}:`, error);
          }
        }
      }
      if (updated) console.info(`Coordinate INRETE Modena salvate in Firestore: ${updated} documenti aggiornati.`);
      return { updated };
    })().finally(() => { persistenceRunning = null; });

    return persistenceRunning;
  }

  function restoreAndPersistRows(rows) {
    restoreRows(rows);
    persistRestoredRows(rows).catch((error) => console.warn("Persistenza coordinate INRETE Modena non completata:", error));
    return rows;
  }

  function installRenderHook() {
    let originalRender = root.renderImpianti;
    try { if (typeof renderImpianti === "function") originalRender = renderImpianti; } catch (_) {}
    if (typeof originalRender !== "function" || originalRender.__inreteModenaCoordinateBackup) return false;

    const patchedRender = function renderImpiantiWithModenaCoordinateBackup() {
      restoreAndPersistRows(getCurrentRows());
      return originalRender.apply(this, arguments);
    };
    patchedRender.__inreteModenaCoordinateBackup = true;
    patchedRender.__originalRender = originalRender;
    root.renderImpianti = patchedRender;
    try { renderImpianti = patchedRender; } catch (_) {}
    return true;
  }

  root.HeraInreteModenaCoordinateBackup = Object.freeze({
    source: SOURCE,
    backupRows: BACKUP_ROWS,
    findBackupRow,
    hasCoordinates,
    readCoordinates,
    restoreRecord,
    restoreRows,
    restoreAndPersistRows,
    persistRestoredRows,
    installRenderHook
  });
  installRenderHook();
})(window);

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

    record.gpsY = backup.lat;
    record.gpsX = backup.lng;
    record.latitudine = backup.lat;
    record.longitudine = backup.lng;
    record.latitude = backup.lat;
    record.longitude = backup.lng;
    record.coordinateStatus = "RESTORED_FROM_BACKUP";
    record.coordinateIssue = "";
    record.coordinateResolvedFrom = SOURCE;
    if (!String(record.coordinateLatitudineOriginale ?? "").trim()) {
      record.coordinateLatitudineOriginale = backup.lat;
    }
    if (!String(record.coordinateLongitudineOriginale ?? "").trim()) {
      record.coordinateLongitudineOriginale = backup.lng;
    }
    return true;
  }

  function restoreRows(rows) {
    if (!Array.isArray(rows)) return rows;
    rows.forEach(restoreRecord);
    return rows;
  }

  function installRenderHook() {
    let originalRender = root.renderImpianti;
    try {
      if (typeof renderImpianti === "function") originalRender = renderImpianti;
    } catch (_) {}
    if (typeof originalRender !== "function" || originalRender.__inreteModenaCoordinateBackup) return false;

    const patchedRender = function renderImpiantiWithModenaCoordinateBackup() {
      try {
        if (typeof currentImpianti !== "undefined") restoreRows(currentImpianti);
      } catch (_) {
        restoreRows(root.currentImpianti);
      }
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
    installRenderHook
  });
  installRenderHook();
})(window);

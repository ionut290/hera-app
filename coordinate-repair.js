(function exposeCoordinateRepair(root) {
  "use strict";

  const ITALY_BOUNDS = Object.freeze({
    minLat: 35,
    maxLat: 48.8,
    minLng: 5,
    maxLng: 20
  });

  function cleanRaw(value) {
    return value == null ? "" : String(value).trim();
  }

  function parseSingle(value) {
    const raw = cleanRaw(value);
    if (!raw) return null;
    const normalized = raw.replace(/\s+/g, "").replace(",", ".");
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

  root.HeraCoordinateRepair = Object.freeze({
    diagnose,
    parseSingle,
    extractPair,
    isValidLatitude,
    isValidLongitude
  });
})(window);

/* Carica la nuova implementazione isolata di “Usa la mia posizione”. */
(function loadCommessaCurrentLocation(root) {
  "use strict";
  if (root.__commessaCurrentLocationLoader) return;
  root.__commessaCurrentLocationLoader = true;

  function load() {
    if (document.querySelector('script[data-commessa-current-location="1"]')) return;
    const script = document.createElement("script");
    script.src = "commessa-current-location.js?v=20260904-1";
    script.async = false;
    script.dataset.commessaCurrentLocation = "1";
    script.onerror = function () {
      console.error("Impossibile caricare il nuovo flusso Usa la mia posizione.");
    };
    (document.head || document.documentElement).appendChild(script);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", load, { once: true });
  } else {
    load();
  }
})(window);

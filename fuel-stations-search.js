(function (root, factory) {
  const api = factory(root);
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.HeraFuelStationSearch = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function (root) {
  "use strict";

  const SAME_ORIGIN_ENDPOINTS = ["/api/fuel-stations/search", "/.netlify/functions/fuel-stations"];
  const OVERPASS_ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.nchc.org.tw/api/interpreter"
  ];

  function timeoutSignal(timeoutMs, externalSignal) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);
    const abort = () => controller.abort();
    externalSignal?.addEventListener?.("abort", abort, { once: true });
    return {
      signal: controller.signal,
      clear() {
        clearTimeout(timeout);
        externalSignal?.removeEventListener?.("abort", abort);
      }
    };
  }

  async function fetchJson(url, options = {}, timeoutMs = 12000, externalSignal) {
    const timer = timeoutSignal(timeoutMs, externalSignal);
    try {
      const response = await root.fetch(url, { ...options, signal: timer.signal });
      const contentType = String(response.headers?.get?.("content-type") || "");
      if (!response.ok) throw Object.assign(new Error(`HTTP ${response.status}`), { status: response.status });
      if (!contentType.includes("json")) throw new Error("Risposta non JSON");
      return await response.json();
    } finally {
      timer.clear();
    }
  }

  function uniqueStations(stations) {
    const seen = new Set();
    return (stations || []).filter((station) => {
      const key = `${Number(station.lat).toFixed(5)}:${Number(station.lon).toFixed(5)}:${station.brandLabel}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  async function searchSameOrigin(position, radiusKm, fuel, distanceFn, signal) {
    const query = new URLSearchParams({ lat: position.lat, lng: position.lng, radius: radiusKm, fuel });
    const errors = [];
    for (const endpoint of SAME_ORIGIN_ENDPOINTS) {
      try {
        const payload = await fetchJson(`${endpoint}?${query}`, { headers: { Accept: "application/json" } }, 12000, signal);
        const stations = Array.isArray(payload.stations)
          ? payload.stations.map((station) => ({
            ...station,
            distance: distanceFn(position.lat, position.lng, Number(station.lat), Number(station.lon))
          }))
          : [];
        if (stations.length || payload.ok) {
          return { stations, source: payload.source || "Servizio distributori", received: payload.received || stations.length };
        }
      } catch (error) {
        if (error?.name === "AbortError" && signal?.aborted) throw error;
        errors.push(`${endpoint}: ${error?.message || error}`);
      }
    }
    throw new Error(errors.join(" | ") || "Servizio hosting non disponibile");
  }

  async function searchNationalCache(position, radiusKm, fuel, distanceFn) {
    const cache = root.HeraFuelNationalCache;
    if (!cache?.findNearby) return null;
    let result = await cache.findNearby(fuel, position, radiusKm, distanceFn);
    if (!result.available) {
      await cache.refresh?.({ force: true });
      result = await cache.findNearby(fuel, position, radiusKm, distanceFn);
    }
    if (!result.available) return null;
    return {
      stations: result.stations,
      source: "Archivio MIMIT salvato",
      received: result.totalCached || result.stations.length
    };
  }

  async function searchMimit(position, radiusKm, fuel, distanceFn, signal) {
    const core = root.HeraFuelStations;
    const payload = await fetchJson(core.MIMIT_API_URL, {
      method: "POST",
      headers: { Accept: "application/json", "Content-Type": "application/json" },
      body: JSON.stringify(core.buildMimitRequest(position.lat, position.lng, radiusKm))
    }, 12000, signal);
    if (payload?.success === false || !Array.isArray(payload?.results)) throw new Error("Risposta MIMIT non valida");
    return {
      stations: core.parseMimitStations(payload.results, fuel, position, distanceFn),
      source: "MIMIT Osservaprezzi",
      received: payload.results.length
    };
  }

  async function searchOverpass(position, radiusKm, fuel, distanceFn, signal) {
    const core = root.HeraFuelStations;
    const query = core.buildQuery(position.lat, position.lng, radiusKm, fuel);
    return Promise.any(OVERPASS_ENDPOINTS.map(async (endpoint) => {
      const payload = await fetchJson(`${endpoint}?data=${encodeURIComponent(query)}`, {
        headers: { Accept: "application/json" }
      }, 18000, signal);
      if (!Array.isArray(payload?.elements)) throw new Error("Risposta OpenStreetMap non valida");
      return {
        stations: core.parseStations(payload.elements, fuel, position, distanceFn)
          .map((station) => ({ ...station, source: "OpenStreetMap" })),
        source: "OpenStreetMap",
        received: payload.elements.length
      };
    }));
  }

  async function search({ position, radiusKm, fuel, distanceFn, signal, onProgress }) {
    const attempts = [
      ["servizio dell’app", () => searchSameOrigin(position, radiusKm, fuel, distanceFn, signal)],
      ["archivio salvato", () => searchNationalCache(position, radiusKm, fuel, distanceFn)],
      ["MIMIT", () => searchMimit(position, radiusKm, fuel, distanceFn, signal)],
      ["OpenStreetMap", () => searchOverpass(position, radiusKm, fuel, distanceFn, signal)]
    ];
    const diagnostics = [];
    let validEmptyResult = null;
    for (const [label, run] of attempts) {
      onProgress?.(label);
      try {
        const result = await run();
        if (!result) continue;
        const stations = uniqueStations(result.stations)
          .filter((station) => Number.isFinite(station.distance) && station.distance <= radiusKm)
          .sort((a, b) => a.distance - b.distance);
        if (stations.length) return { ...result, stations, diagnostics };
        validEmptyResult = { ...result, stations: [], diagnostics };
        diagnostics.push(`${label}: nessun risultato compatibile`);
      } catch (error) {
        if (error?.name === "AbortError" && signal?.aborted) throw error;
        diagnostics.push(`${label}: ${error?.message || error}`);
      }
    }
    if (validEmptyResult) return validEmptyResult;
    throw Object.assign(new Error("Tutte le fonti distributori non disponibili"), { diagnostics });
  }

  return { SAME_ORIGIN_ENDPOINTS, OVERPASS_ENDPOINTS, uniqueStations, search };
});

(function (root, factory) {
  const api = factory(root);
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.HeraFuelNationalCache = api;
  if (root?.document) api.scheduleRefresh();
})(typeof globalThis !== "undefined" ? globalThis : this, function (root) {
  "use strict";

  const DB_NAME = "hera-fuel-national-cache";
  const DB_VERSION = 1;
  const STORE_NAME = "snapshots";
  const SNAPSHOT_KEY = "mimit-italia";
  const ENDPOINT = "/api/fuel-stations-italy";
  const CACHE_TTL_MS = 24 * 60 * 60 * 1000;
  const MAX_STALE_MS = 7 * 24 * 60 * 60 * 1000;
  let memorySnapshot = null;
  let refreshPromise = null;
  let refreshScheduled = false;

  function snapshotAge(snapshot, now = Date.now()) {
    return snapshot?.updatedAt ? Math.max(0, now - Number(snapshot.updatedAt)) : Infinity;
  }

  function normalizeSnapshot(payload) {
    if (!payload || !Array.isArray(payload.stations)) throw new Error("Archivio nazionale distributori non valido");
    return {
      key: SNAPSHOT_KEY,
      updatedAt: Number(payload.updatedAt) || Date.now(),
      extractionDate: String(payload.extractionDate || ""),
      stations: payload.stations
    };
  }

  function openDatabase() {
    return new Promise((resolve, reject) => {
      if (!root?.indexedDB) return reject(new Error("IndexedDB non disponibile"));
      const request = root.indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains(STORE_NAME)) database.createObjectStore(STORE_NAME, { keyPath: "key" });
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Apertura archivio distributori non riuscita"));
    });
  }

  async function readSnapshot() {
    if (memorySnapshot) return memorySnapshot;
    const database = await openDatabase();
    try {
      memorySnapshot = await new Promise((resolve, reject) => {
        const request = database.transaction(STORE_NAME, "readonly").objectStore(STORE_NAME).get(SNAPSHOT_KEY);
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error || new Error("Lettura archivio distributori non riuscita"));
      });
      return memorySnapshot;
    } finally {
      database.close();
    }
  }

  async function writeSnapshot(snapshot) {
    const database = await openDatabase();
    try {
      await new Promise((resolve, reject) => {
        const request = database.transaction(STORE_NAME, "readwrite").objectStore(STORE_NAME).put(snapshot);
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error || new Error("Salvataggio archivio distributori non riuscito"));
      });
      memorySnapshot = snapshot;
      return snapshot;
    } finally {
      database.close();
    }
  }

  async function refresh(options = {}) {
    if (refreshPromise) return refreshPromise;
    refreshPromise = (async () => {
      let existing = null;
      try {
        existing = await readSnapshot();
      } catch (error) {
        console.warn("[Distributori] archivio locale non leggibile:", error);
      }
      if (!options.force && snapshotAge(existing) < CACHE_TTL_MS) return { snapshot: existing, refreshed: false };
      if (!root?.fetch || root?.navigator?.onLine === false) return { snapshot: existing, refreshed: false };
      if (root?.navigator?.connection?.saveData) return { snapshot: existing, refreshed: false };

      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 45000);
      try {
        const response = await root.fetch(ENDPOINT, {
          method: "GET",
          headers: { Accept: "application/json" },
          signal: controller.signal
        });
        if (!response.ok) throw Object.assign(new Error(`Archivio distributori HTTP ${response.status}`), { status: response.status });
        const snapshot = normalizeSnapshot(await response.json());
        if (!snapshot.stations.length) throw new Error("Archivio nazionale distributori vuoto");
        await writeSnapshot(snapshot);
        root?.dispatchEvent?.(new CustomEvent("hera:fuel-national-cache-ready", {
          detail: { count: snapshot.stations.length, updatedAt: snapshot.updatedAt }
        }));
        console.info("[Distributori] archivio MIMIT nazionale salvato", { count: snapshot.stations.length });
        return { snapshot, refreshed: true };
      } finally {
        clearTimeout(timeout);
      }
    })().catch((error) => {
      console.warn("[Distributori] aggiornamento archivio nazionale non riuscito, resta attiva la ricerca online:", error);
      return { snapshot: memorySnapshot, refreshed: false, error };
    }).finally(() => {
      refreshPromise = null;
    });
    return refreshPromise;
  }

  async function findNearby(fuel, origin, radiusKm, distanceFn) {
    if (fuel === "electric") return { stations: [], available: false };
    let snapshot;
    try {
      snapshot = await readSnapshot();
    } catch (error) {
      return { stations: [], available: false };
    }
    if (!snapshot || snapshotAge(snapshot) > MAX_STALE_MS) return { stations: [], available: false };
    const parser = root?.HeraFuelStations?.parseMimitStations;
    if (typeof parser !== "function") return { stations: [], available: false };
    const stations = parser(snapshot.stations, fuel, origin, distanceFn)
      .filter((station) => station.distance <= radiusKm);
    return {
      stations,
      available: true,
      updatedAt: snapshot.updatedAt,
      extractionDate: snapshot.extractionDate,
      totalCached: snapshot.stations.length
    };
  }

  function scheduleRefresh() {
    if (refreshScheduled) return;
    refreshScheduled = true;
    const start = () => {
      const run = () => refresh().catch(() => {});
      if (typeof root.requestIdleCallback === "function") root.requestIdleCallback(run, { timeout: 15000 });
      else setTimeout(run, 12000);
    };
    if (root.document.readyState === "complete") start();
    else root.addEventListener("load", start, { once: true });
    root.addEventListener("online", () => refresh().catch(() => {}));
  }

  return {
    ENDPOINT,
    CACHE_TTL_MS,
    MAX_STALE_MS,
    snapshotAge,
    normalizeSnapshot,
    readSnapshot,
    refresh,
    findNearby,
    scheduleRefresh
  };
});

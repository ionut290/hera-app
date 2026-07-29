(function installWhazzupPreloadCache() {
  "use strict";

  const CACHE_VERSION = 1;
  const CACHE_KEY = "hera_whazzup_preload_cache_v1";
  const memoryCache = new Map();
  let preloadTimer = null;

  function safeString(value) {
    return value == null ? "" : String(value);
  }

  function getImpiantoKey(impianto) {
    return safeString(
      impianto?.id ||
      impianto?.key ||
      impianto?.impiantoId ||
      impianto?.idSap ||
      impianto?.sap ||
      impianto?.denominazione ||
      impianto?.nome
    ).trim();
  }

  function buildFingerprint(impianto) {
    const source = {
      id: getImpiantoKey(impianto),
      idSap: impianto?.idSap || impianto?.sap || impianto?.ID_SAP || "",
      nome: impianto?.denominazione || impianto?.nome || impianto?.impianto || "",
      comune: impianto?.comune || impianto?.Comune || "",
      indirizzo: impianto?.indirizzo || impianto?.via || impianto?.descrizioneVia || "",
      coordinate: impianto?.coordinate || impianto?.coords || impianto?.latLng || "",
      lat: impianto?.lat || impianto?.latitude || "",
      lng: impianto?.lng || impianto?.lon || impianto?.longitude || "",
      commessaId: impianto?.commessaId || impianto?.projectId || "",
      commessa: impianto?.commessa || impianto?.commessaName || "",
      lavorazione: impianto?.lavorazione || impianto?.tipoLavoro || "",
      nota: impianto?.nota || ""
    };
    return JSON.stringify(source);
  }

  function readPersistentCache() {
    try {
      const raw = localStorage.getItem(CACHE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (!parsed || parsed.version !== CACHE_VERSION || !parsed.items) return;
      Object.entries(parsed.items).forEach(([key, value]) => {
        if (value && typeof value === "object") memoryCache.set(key, value);
      });
    } catch (error) {
      console.warn("Cache Whazzup non leggibile:", error);
    }
  }

  function persistCache() {
    try {
      const items = {};
      memoryCache.forEach((value, key) => {
        items[key] = value;
      });
      localStorage.setItem(CACHE_KEY, JSON.stringify({
        version: CACHE_VERSION,
        updatedAt: new Date().toISOString(),
        items
      }));
    } catch (error) {
      console.warn("Cache Whazzup non salvata:", error);
    }
  }

  function normalizePreparedValue(value) {
    if (!value) return null;
    if (typeof value === "string") return { text: value };
    if (typeof value === "object") {
      if (typeof value.text === "string") return { ...value, text: value.text };
      if (typeof value.message === "string") return { ...value, text: value.message };
      if (typeof value.url === "string") return { ...value, url: value.url };
    }
    return null;
  }

  function resolveBuilder() {
    const candidates = [
      "buildWhatsAppMessage",
      "buildWhazzupMessage",
      "createWhatsAppMessage",
      "createWhazzupMessage",
      "prepareWhatsAppMessage",
      "prepareWhazzupMessage",
      "getWhatsAppMessage",
      "getWhazzupMessage"
    ];
    for (const name of candidates) {
      if (typeof window[name] === "function") return window[name];
    }
    return null;
  }

  function prepareOne(impianto, options = {}) {
    const key = getImpiantoKey(impianto);
    if (!key) return null;
    const fingerprint = buildFingerprint(impianto);
    const existing = memoryCache.get(key);
    if (existing?.fingerprint === fingerprint && existing?.prepared) return existing.prepared;

    const builder = options.builder || resolveBuilder();
    if (typeof builder !== "function") return null;

    try {
      const prepared = normalizePreparedValue(builder(impianto, {
        preload: true,
        dynamicValues: false,
        skipOpen: true
      }));
      if (!prepared) return null;
      memoryCache.set(key, {
        fingerprint,
        prepared,
        preparedAt: Date.now()
      });
      return prepared;
    } catch (error) {
      console.debug("Precaricamento Whazzup non disponibile per", key, error);
      return null;
    }
  }

  function extractImpiantiFromWindow() {
    const candidates = [
      window.impianti,
      window.allImpianti,
      window.currentImpianti,
      window.impiantiList,
      window.appState?.impianti,
      window.state?.impianti
    ];
    for (const value of candidates) {
      if (Array.isArray(value)) return value;
      if (value instanceof Map) return Array.from(value.values());
      if (value && typeof value === "object") {
        const list = Object.values(value).filter((item) => item && typeof item === "object");
        if (list.length) return list;
      }
    }
    return [];
  }

  function preload(impianti, options = {}) {
    const list = Array.isArray(impianti) ? impianti : extractImpiantiFromWindow();
    if (!list.length) return 0;
    let preparedCount = 0;
    list.forEach((impianto) => {
      if (prepareOne(impianto, options)) preparedCount += 1;
    });
    if (preparedCount) persistCache();
    return preparedCount;
  }

  function schedulePreload(impianti, options = {}) {
    clearTimeout(preloadTimer);
    preloadTimer = setTimeout(() => {
      const run = () => preload(impianti, options);
      if (typeof requestIdleCallback === "function") {
        requestIdleCallback(run, { timeout: 1200 });
      } else {
        run();
      }
    }, options.delay ?? 60);
  }

  function getPrepared(impianto) {
    const key = getImpiantoKey(impianto);
    if (!key) return null;
    const entry = memoryCache.get(key);
    if (!entry || entry.fingerprint !== buildFingerprint(impianto)) return null;
    return entry.prepared || null;
  }

  function invalidate(impiantoOrKey) {
    const key = typeof impiantoOrKey === "string" ? impiantoOrKey : getImpiantoKey(impiantoOrKey);
    if (!key) return false;
    const deleted = memoryCache.delete(key);
    if (deleted) persistCache();
    return deleted;
  }

  function clear() {
    memoryCache.clear();
    try { localStorage.removeItem(CACHE_KEY); } catch (_) {}
  }

  readPersistentCache();

  window.HeraWhazzupPreload = {
    version: CACHE_VERSION,
    preload,
    schedulePreload,
    prepareOne,
    getPrepared,
    invalidate,
    clear
  };

  window.addEventListener("load", () => schedulePreload());
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") schedulePreload();
  });
  document.addEventListener("hera:impianti-loaded", (event) => {
    schedulePreload(event?.detail?.impianti || event?.detail || []);
  });
  document.addEventListener("hera:impianto-updated", (event) => {
    const impianto = event?.detail?.impianto || event?.detail;
    if (impianto) {
      invalidate(impianto);
      schedulePreload([impianto]);
    }
  });
})();

(() => {
  "use strict";

  const PAGE_ID = "verde-bologna-page";
  const CATEGORY_ID = "verde-bologna-operativo-category";
  const PARKS_DATASET_ID = "carta-tecnica-comunale-toponimi-parchi-e-giardini";
  const MANAGED_AREAS_DATASET_ID = "un_gest";
  const QUARTIERI_DATASET_ID = "quartieri-di-bologna";
  const API_ROOT = "https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets";
  const MOBILE_QUERY = "(max-width: 760px)";
  const STYLE_ID = "verde-bologna-parchi-mobile-style";
  const FILTERS_ID = "verde-bologna-parchi-quartieri";
  const LIST_ID = "verde-bologna-parchi-list";
  const SHEET_ID = "verde-bologna-parchi-sheet";
  const CACHE_TTL_MS = 10 * 60 * 1000;
  const PAGE_SIZE = 100;
  const FETCH_TIMEOUT_MS = 8000;
  const BASE_CACHE_WAIT_MS = 1500;
  const BASE_CACHE_POLL_MS = 150;
  const SEARCH_DEBOUNCE_MS = 180;
  const LIST_RENDER_LIMIT = 60;
  const LABEL_MARKER_LIMIT = 80;
  const MANAGED_BOUNDARY_RADIUS_METERS = 500;
  const managedBoundaryCache = new Map();
  const managedBoundaryRequests = new Map();
  const OFFICIAL_QUARTIERI = [
    "Borgo Panigale - Reno",
    "Navile",
    "Porto - Saragozza",
    "San Donato - San Vitale",
    "Santo Stefano",
    "Savena"
  ];

  const state = {
    map: null,
    layer: null,
    boundaryLayer: null,
    parks: [],
    quartieri: [],
    filtered: [],
    loading: false,
    loaded: false,
    complete: false,
    activeQuartiere: "",
    userPosition: null,
    locationRequested: false,
    listOpenRecord: null,
    installDone: false,
    markerRenderer: null,
    mapEventsBound: false,
    markerRefreshTimer: 0,
    filterTimer: 0,
    renderSignature: "",
    generation: 0,
    abortControllers: new Set(),
    active: false
  };

  const $ = (id) => document.getElementById(id);
  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[char]));

  function mobileActive() {
    return window.matchMedia(MOBILE_QUERY).matches;
  }

  function parksActive() {
    return mobileActive() && $(CATEGORY_ID)?.value === PARKS_DATASET_ID && $(PAGE_ID)?.classList.contains("is-category-open") && !$(PAGE_ID)?.classList.contains("hidden");
  }

  function normalizeText(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLocaleLowerCase("it-IT")
      .trim();
  }

  function fieldValue(record, names) {
    const entries = Object.entries(record || {});
    const normalized = new Map(entries.map(([key, value]) => [normalizeText(key).replace(/[^a-z0-9]/g, ""), value]));
    for (const name of names) {
      const key = normalizeText(name).replace(/[^a-z0-9]/g, "");
      const value = normalized.get(key);
      if (value !== undefined && value !== null && String(value).trim() !== "") return value;
    }
    return "";
  }

  function parkCodvia(record) {
    if (record?.__vbCodvia !== undefined) return record.__vbCodvia;
    const value = fieldValue(record, ["codvia", "cod_via", "codice via"]);
    return value === "" ? "" : String(value).trim();
  }

  function parkFallbackCode(record) {
    return String(fieldValue(record, ["cod_ogg", "codogg", "idtopon", "id", "objectid"]) || "").trim();
  }

  function parkName(record) {
    if (record?.__vbName !== undefined) return record.__vbName;
    return String(fieldValue(record, [
      "nomevia", "nome_via", "nome via", "completo", "porzione", "denominazione", "nome", "name", "toponimo"
    ]) || "Parco / giardino").trim();
  }

  function parseGeometryValue(value) {
    if (!value) return null;
    if (typeof value === "string") {
      try { return parseGeometryValue(JSON.parse(value)); } catch (_) { return null; }
    }
    if (value.type === "Feature" && value.geometry) return parseGeometryValue(value.geometry);
    if (value.geometry) return parseGeometryValue(value.geometry);
    if (value.type && Array.isArray(value.coordinates)) return value;
    if (Number.isFinite(Number(value.lat)) && Number.isFinite(Number(value.lon ?? value.lng))) {
      return { type: "Point", coordinates: [Number(value.lon ?? value.lng), Number(value.lat)] };
    }
    return null;
  }

  function geometryPriority(geometry) {
    if (geometry?.type === "Polygon" || geometry?.type === "MultiPolygon") return 3;
    if (geometry?.type === "LineString" || geometry?.type === "MultiLineString") return 2;
    if (geometry?.type === "Point" || geometry?.type === "MultiPoint") return 1;
    return 0;
  }

  function geometryOf(record) {
    let selected = null;
    for (const [key, value] of Object.entries(record || {})) {
      const normalized = normalizeText(key).replace(/[^a-z0-9]/g, "");
      if (!normalized.includes("geoshape") && normalized !== "geometry" && normalized !== "geom" && !normalized.includes("geopoint")) continue;
      const geometry = parseGeometryValue(value);
      if (geometryPriority(geometry) > geometryPriority(selected)) selected = geometry;
    }
    if (selected) return selected;
    const point = fieldValue(record, ["geo_point_2d", "geopoint", "geo point"]);
    const parsed = parseGeometryValue(point);
    if (parsed) return parsed;
    const lat = Number(fieldValue(record, ["lat", "latitude", "y"]));
    const lon = Number(fieldValue(record, ["lon", "lng", "longitude", "x"]));
    if (Number.isFinite(lat) && Number.isFinite(lon)) return { type: "Point", coordinates: [lon, lat] };
    return null;
  }

  function flattenCoordinates(coords, output = []) {
    if (!Array.isArray(coords)) return output;
    if (coords.length >= 2 && Number.isFinite(Number(coords[0])) && Number.isFinite(Number(coords[1]))) {
      output.push([Number(coords[0]), Number(coords[1])]);
      return output;
    }
    coords.forEach((item) => flattenCoordinates(item, output));
    return output;
  }

  function centerOf(record) {
    if (record?.__vbCenter !== undefined) return record.__vbCenter;
    const geometry = geometryOf(record);
    if (!geometry) return null;
    const points = flattenCoordinates(geometry.coordinates).filter(([lon, lat]) => Math.abs(lon) <= 180 && Math.abs(lat) <= 90);
    if (!points.length) return null;
    if (geometry.type === "Point") return { lon: points[0][0], lat: points[0][1] };
    const sum = points.reduce((acc, [lon, lat]) => ({ lon: acc.lon + lon, lat: acc.lat + lat }), { lon: 0, lat: 0 });
    return { lon: sum.lon / points.length, lat: sum.lat / points.length };
  }

  function pointInRing(lon, lat, ring) {
    let inside = false;
    for (let i = 0, j = ring.length - 1; i < ring.length; j = i++) {
      const xi = Number(ring[i]?.[0]);
      const yi = Number(ring[i]?.[1]);
      const xj = Number(ring[j]?.[0]);
      const yj = Number(ring[j]?.[1]);
      if (![xi, yi, xj, yj].every(Number.isFinite)) continue;
      const intersects = ((yi > lat) !== (yj > lat)) && (lon < ((xj - xi) * (lat - yi)) / ((yj - yi) || Number.EPSILON) + xi);
      if (intersects) inside = !inside;
    }
    return inside;
  }

  function pointInPolygon(lon, lat, polygon) {
    if (!Array.isArray(polygon) || !polygon.length) return false;
    if (!pointInRing(lon, lat, polygon[0])) return false;
    for (let i = 1; i < polygon.length; i += 1) {
      if (pointInRing(lon, lat, polygon[i])) return false;
    }
    return true;
  }

  function pointInGeometry(lon, lat, geometry) {
    if (!geometry) return false;
    if (geometry.type === "Polygon") return pointInPolygon(lon, lat, geometry.coordinates);
    if (geometry.type === "MultiPolygon") return geometry.coordinates.some((polygon) => pointInPolygon(lon, lat, polygon));
    return false;
  }

  function quartiereName(record) {
    return String(fieldValue(record, ["quartiere", "nomequartiere", "nome_quartiere", "nome", "denominazione"]) || "").trim();
  }

  function assignQuartiere(record) {
    const officialValue = String(fieldValue(record, ["quartiere", "nomequartiere", "nome_quartiere"]) || "").trim();
    if (officialValue) return officialValue;
    const center = centerOf(record);
    if (!center) return "";
    const hit = state.quartieri.find((quartiere) => pointInGeometry(center.lon, center.lat, geometryOf(quartiere)));
    return hit ? quartiereName(hit) : "";
  }

  function haversineMeters(a, b) {
    if (!a || !b) return Number.POSITIVE_INFINITY;
    const rad = (degrees) => degrees * Math.PI / 180;
    const earth = 6371000;
    const dLat = rad(b.lat - a.lat);
    const dLon = rad(b.lon - a.lon);
    const lat1 = rad(a.lat);
    const lat2 = rad(b.lat);
    const h = Math.sin(dLat / 2) ** 2 + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLon / 2) ** 2;
    return 2 * earth * Math.asin(Math.min(1, Math.sqrt(h)));
  }

  function recordDistance(record) {
    const center = centerOf(record);
    if (!center || !state.userPosition) return Number.POSITIVE_INFINITY;
    return haversineMeters(state.userPosition, center);
  }

  function meaningfulParkWords(value) {
    const stopWords = new Set([
      "area", "parco", "parchi", "pco", "giardino", "giardini", "verde", "del", "della", "delle",
      "dei", "degli", "di", "da", "il", "lo", "la", "i", "gli", "le", "e"
    ]);
    return normalizeText(value)
      .replace(/[^a-z0-9]+/g, " ")
      .split(/\s+/)
      .filter((word) => word.length > 2 && !stopWords.has(word));
  }

  function escapeSearch(value) {
    return String(value || "").replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  }

  function areaGeometryOf(record) {
    const managed = parseGeometryValue(record?.__vbBoundaryGeometry);
    if (managed?.type === "Polygon" || managed?.type === "MultiPolygon") return managed;
    const native = geometryOf(record);
    return native?.type === "Polygon" || native?.type === "MultiPolygon" ? native : null;
  }

  function managedRecordName(record) {
    return String(fieldValue(record, ["nome", "nome_ug", "denominazione", "ubicazione"]) || "").trim();
  }

  function managedRecordCenter(record) {
    const point = parseGeometryValue(fieldValue(record, ["geo_point_2d", "geopoint", "geo point"]));
    if (point?.type === "Point" && Array.isArray(point.coordinates)) {
      const [lon, lat] = point.coordinates.map(Number);
      if (Number.isFinite(lat) && Number.isFinite(lon)) return { lat, lon };
    }
    const geometry = parseGeometryValue(record?.geo_shape ?? record?.geometry);
    if (!geometry) return null;
    const points = flattenCoordinates(geometry.coordinates).filter(([lon, lat]) => Math.abs(lon) <= 180 && Math.abs(lat) <= 90);
    if (!points.length) return null;
    const sum = points.reduce((acc, [lon, lat]) => ({ lon: acc.lon + lon, lat: acc.lat + lat }), { lon: 0, lat: 0 });
    return { lon: sum.lon / points.length, lat: sum.lat / points.length };
  }

  function boundaryCacheKey(record) {
    const center = centerOf(record);
    return `${normalizeText(parkName(record))}:${center ? `${center.lat.toFixed(5)},${center.lon.toFixed(5)}` : parkCodvia(record)}`;
  }

  function selectManagedBoundary(parkRecord, candidates) {
    const parkCenter = centerOf(parkRecord);
    const parkWords = meaningfulParkWords(parkName(parkRecord));
    const unique = new Map();
    (candidates || []).forEach((candidate) => {
      const geometry = parseGeometryValue(candidate?.geo_shape ?? candidate?.geometry);
      if (geometry?.type !== "Polygon" && geometry?.type !== "MultiPolygon") return;
      const name = managedRecordName(candidate);
      const center = managedRecordCenter(candidate);
      const candidateWords = meaningfulParkWords(name);
      const overlap = parkWords.filter((word) => candidateWords.includes(word)).length;
      const coverage = parkWords.length ? overlap / parkWords.length : 0;
      const precision = candidateWords.length ? overlap / candidateWords.length : 0;
      const distance = haversineMeters(parkCenter, center);
      const score = coverage * 100 + precision * 40 - Math.min(distance, 5000) / 50;
      const key = JSON.stringify(geometry);
      const previous = unique.get(key);
      if (!previous || score > previous.score) unique.set(key, { geometry, coverage, distance, score });
    });
    const ranked = [...unique.values()].sort((a, b) => b.score - a.score);
    const best = ranked[0];
    if (!best) return null;
    if ((best.coverage >= 0.5 && best.distance <= 2500) || best.distance <= 180) return best.geometry;
    return null;
  }

  async function fetchManagedRecords(where, controller) {
    const params = new URLSearchParams({ limit: "100", where });
    const response = await fetch(`${API_ROOT}/${MANAGED_AREAS_DATASET_ID}/records?${params}`, {
      headers: { Accept: "application/json" },
      signal: controller?.signal
    });
    if (!response.ok) throw new Error(`Confini comunali non disponibili (${response.status}).`);
    const payload = await response.json();
    return Array.isArray(payload?.results) ? payload.results : [];
  }

  async function fetchManagedBoundary(record) {
    const center = centerOf(record);
    const words = meaningfulParkWords(parkName(record));
    const controller = typeof AbortController === "function" ? new AbortController() : null;
    if (controller) state.abortControllers.add(controller);
    const timer = window.setTimeout(() => controller?.abort(), FETCH_TIMEOUT_MS);
    try {
      let candidates = [];
      if (words.length) {
        const phrase = words.slice(-3).join(" ");
        candidates = await fetchManagedRecords(`search("${escapeSearch(phrase)}")`, controller);
        const namedBoundary = selectManagedBoundary(record, candidates);
        if (namedBoundary) return namedBoundary;
      }
      if (center) {
        const where = `within_distance(geo_point_2d, geom'POINT(${center.lon} ${center.lat})', ${MANAGED_BOUNDARY_RADIUS_METERS}m)`;
        const nearby = await fetchManagedRecords(where, controller);
        candidates = [...candidates, ...nearby];
      }
      return selectManagedBoundary(record, candidates);
    } finally {
      window.clearTimeout(timer);
      if (controller) state.abortControllers.delete(controller);
    }
  }

  async function loadManagedBoundary(record) {
    if (!record || areaGeometryOf(record)) {
      if (record) record.__vbBoundaryStatus = "available";
      return;
    }
    const cacheKey = boundaryCacheKey(record);
    const cached = managedBoundaryCache.get(cacheKey) ?? readSessionCache(`varga-verde-bologna:park-boundary:${cacheKey}`);
    if (cached && typeof cached === "object" && "found" in cached) {
      record.__vbBoundaryGeometry = cached.found ? cached.geometry : null;
      record.__vbBoundaryStatus = cached.found ? "available" : "unavailable";
      if (state.listOpenRecord === record) renderDetailSheet(record);
      return;
    }
    if (managedBoundaryRequests.has(cacheKey)) {
      await managedBoundaryRequests.get(cacheKey);
      const shared = managedBoundaryCache.get(cacheKey);
      record.__vbBoundaryGeometry = shared?.found ? shared.geometry : null;
      record.__vbBoundaryStatus = shared?.found ? "available" : "unavailable";
      if (state.listOpenRecord === record) renderDetailSheet(record);
      return;
    }
    record.__vbBoundaryStatus = "loading";
    if (state.listOpenRecord === record) renderDetailSheet(record);
    const request = fetchManagedBoundary(record)
      .then((geometry) => {
        const result = { found: Boolean(geometry), geometry: geometry || null };
        managedBoundaryCache.set(cacheKey, result);
        writeSessionCache(`varga-verde-bologna:park-boundary:${cacheKey}`, result);
        record.__vbBoundaryGeometry = geometry || null;
        record.__vbBoundaryStatus = geometry ? "available" : "unavailable";
      })
      .catch((error) => {
        if (error?.name === "AbortError") record.__vbBoundaryStatus = "unknown";
        else record.__vbBoundaryStatus = "error";
      })
      .finally(() => managedBoundaryRequests.delete(cacheKey));
    managedBoundaryRequests.set(cacheKey, request);
    await request;
    if (state.listOpenRecord === record) renderDetailSheet(record);
  }

  function prepareRecord(record) {
    const codvia = parkCodvia(record);
    const name = parkName(record);
    const quartiere = assignQuartiere(record);
    const center = centerOf(record);
    return {
      ...record,
      __vbCodvia: codvia,
      __vbName: name,
      __vbQuartiere: quartiere,
      __vbCenter: center,
      __vbSearch: normalizeText(`${codvia} ${name}`)
    };
  }

  function readSessionCache(key) {
    try {
      const raw = sessionStorage.getItem(key);
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!parsed || Date.now() - Number(parsed.savedAt || 0) > CACHE_TTL_MS) return null;
      return parsed.data;
    } catch (_) { return null; }
  }

  function writeSessionCache(key, data) {
    try { sessionStorage.setItem(key, JSON.stringify({ savedAt: Date.now(), data })); } catch (_) {}
  }

  function removeLegacyQuartieriCache() {
    try { sessionStorage.removeItem(`varga-verde-bologna:all:${QUARTIERI_DATASET_ID}`); } catch (_) {}
  }

  function readBasePageCache() {
    try {
      const key = `varga-verde-bologna:${PARKS_DATASET_ID}:0:plain:`;
      const raw = sessionStorage.getItem(key);
      if (!raw) return null;
      const cached = JSON.parse(raw);
      if (!cached || Date.now() - Number(cached.savedAt || 0) > CACHE_TTL_MS) return null;
      const records = Array.isArray(cached.payload?.results) ? cached.payload.results : null;
      if (!records) return null;
      return { records, total: Number(cached.payload?.total_count ?? records.length) || records.length };
    } catch (_) { return null; }
  }

  function delay(ms) {
    return new Promise((resolve) => window.setTimeout(resolve, ms));
  }

  async function waitForBasePageCache() {
    const deadline = Date.now() + BASE_CACHE_WAIT_MS;
    do {
      const cached = readBasePageCache();
      if (cached) return cached;
      await delay(BASE_CACHE_POLL_MS);
    } while (Date.now() < deadline);
    return null;
  }

  async function fetchRecordsPage(datasetId, offset) {
    const params = new URLSearchParams({ limit: String(PAGE_SIZE), offset: String(offset) });
    const controller = typeof AbortController === "function" ? new AbortController() : null;
    if (controller) state.abortControllers.add(controller);
    const timer = window.setTimeout(() => controller?.abort(), FETCH_TIMEOUT_MS);
    try {
      const response = await fetch(`${API_ROOT}/${encodeURIComponent(datasetId)}/records?${params}`, {
        headers: { Accept: "application/json" },
        signal: controller?.signal
      });
      if (!response.ok) throw new Error(`Open Data Comune di Bologna non disponibili (${response.status}).`);
      const payload = await response.json();
      const records = Array.isArray(payload?.results) ? payload.results : [];
      return { records, total: Number(payload?.total_count ?? records.length) || records.length };
    } catch (error) {
      if (error?.name === "AbortError") throw new Error("Il dataset del Comune sta impiegando troppo tempo a rispondere.");
      throw error;
    } finally {
      window.clearTimeout(timer);
      if (controller) state.abortControllers.delete(controller);
    }
  }

  function publishParkRecords(records, { complete = false } = {}) {
    state.parks = records.map(prepareRecord);
    state.loaded = true;
    state.complete = complete;
    state.loading = false;
    state.renderSignature = "";
    renderQuartiereFilters();
    applyLiveFilter();
  }

  function injectStyle() {
    if ($(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
      .verde-bologna-parks-tools,.verde-bologna-parks-list,.verde-bologna-parks-sheet{display:none}
      .verde-bologna-page.vb-parks-advanced .verde-bologna-marker-wrap{display:none!important}
      @media ${MOBILE_QUERY}{
        .verde-bologna-page.vb-parks-advanced .verde-bologna-code-marker-wrap{display:none!important}
        .verde-bologna-page.vb-parks-advanced #verde-bologna-results,
        .verde-bologna-page.vb-parks-advanced #verde-bologna-load-more{display:none!important}
        .verde-bologna-page.vb-parks-advanced .verde-bologna-search{margin-bottom:6px!important}
        .verde-bologna-parks-tools{display:none;margin:0;padding:0 8px 6px}
        .verde-bologna-page.vb-parks-advanced .verde-bologna-parks-tools{display:block}
        .verde-bologna-parks-tools-title{display:none}
        .verde-bologna-parks-chips{display:flex;gap:6px;overflow-x:auto;padding:2px 0 5px;scrollbar-width:none}
        .verde-bologna-parks-chips::-webkit-scrollbar{display:none}
        .verde-bologna-parks-chip{flex:0 0 auto;min-height:36px;padding:7px 10px;border:1px solid #c5d3e1;border-radius:999px;background:#f5f8fb;color:#244766;font-size:.72rem;font-weight:900}
        .verde-bologna-parks-chip.is-active{border-color:#12623a;background:#12623a;color:#fff}
        .verde-bologna-parks-list{display:grid;gap:7px;margin-top:10px}
        .verde-bologna-parks-list-status{margin:0;padding:8px 10px;border-radius:9px;background:#edf4fb;color:#355777;font-size:.75rem}
        .verde-bologna-parks-row{display:grid;grid-template-columns:minmax(74px,.32fr) minmax(0,1fr) auto;gap:9px;align-items:center;width:100%;padding:10px;border:1px solid #d6e1eb;border-radius:12px;background:#fff;text-align:left;box-shadow:0 3px 10px rgba(26,55,91,.06)}
        .verde-bologna-parks-code{display:inline-flex;align-items:center;justify-content:center;min-height:34px;padding:6px 8px;border-radius:10px;background:#12623a;color:#fff;font-size:.78rem;font-weight:900;white-space:nowrap}
        .verde-bologna-parks-name{min-width:0;color:#10264a;font-size:.88rem;font-weight:850;line-height:1.25;overflow:hidden;text-overflow:ellipsis}
        .verde-bologna-parks-open{font-size:1.2rem;color:#4f6882}
        .vb-park-live-wrap{background:transparent!important;border:0!important}
        .vb-park-live-code{display:flex;align-items:center;justify-content:center;min-width:42px;height:29px;padding:0 7px;border:2px solid #fff;border-radius:15px;background:#12623a;color:#fff;box-shadow:0 2px 7px rgba(0,0,0,.4);font-size:.72rem;font-weight:900;white-space:nowrap}
        .vb-park-live-code.is-fallback{background:#58697b}
        .verde-bologna-parks-sheet{position:fixed;inset:0;z-index:13050;overflow:auto;background:#f1f6fb;padding:max(10px,env(safe-area-inset-top)) 10px max(16px,env(safe-area-inset-bottom))}
        .verde-bologna-parks-sheet.is-open{display:block}
        .verde-bologna-parks-sheet-head{position:sticky;top:0;z-index:2;display:flex;gap:9px;align-items:center;margin:-10px -10px 10px;padding:max(10px,env(safe-area-inset-top)) 10px 10px;background:rgba(255,255,255,.97);border-bottom:1px solid #d9e3ef}
        .verde-bologna-parks-sheet-head h2{margin:0;min-width:0;flex:1;color:#10264a;font-size:1.05rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
        .verde-bologna-parks-sheet-summary{display:grid;grid-template-columns:auto 1fr;gap:8px;margin-bottom:10px;padding:12px;border-radius:14px;background:#fff}
        .verde-bologna-parks-sheet-summary strong{color:#12623a}.verde-bologna-parks-sheet-summary span{color:#526b84}
        .verde-bologna-parks-sheet-actions{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px}
        .verde-bologna-parks-sheet-actions .btn{min-height:44px}
        .verde-bologna-parks-fields{display:grid;gap:7px}
        .verde-bologna-parks-field{padding:10px;border-radius:11px;background:#fff;border:1px solid #dce6ef}
        .verde-bologna-parks-field span{display:block;margin-bottom:3px;color:#617990;font-size:.68rem;font-weight:900;text-transform:uppercase;letter-spacing:.03em}
        .verde-bologna-parks-field strong,.verde-bologna-parks-field pre{display:block;margin:0;color:#203e59;font-size:.82rem;line-height:1.35;white-space:pre-wrap;overflow-wrap:anywhere;font-family:inherit}
      }
    `;
    document.head.appendChild(style);
  }

  function ensureUi() {
    const page = $(PAGE_ID);
    const searchForm = $("verde-bologna-search-form");
    const mapCard = $("verde-bologna-map-card");
    if (!page || !searchForm || !mapCard) return false;

    if (!$(FILTERS_ID)) {
      const tools = document.createElement("section");
      tools.id = FILTERS_ID;
      tools.className = "verde-bologna-parks-tools";
      tools.innerHTML = `<p class="verde-bologna-parks-tools-title">Filtra per quartiere</p><div class="verde-bologna-parks-chips" role="group" aria-label="Filtri quartiere"></div>`;
      searchForm.insertAdjacentElement("afterend", tools);
    }

    if (!$(LIST_ID)) {
      const list = document.createElement("section");
      list.id = LIST_ID;
      list.className = "verde-bologna-parks-list";
      list.setAttribute("aria-live", "polite");
      mapCard.insertAdjacentElement("afterend", list);
    }

    if (!$(SHEET_ID)) {
      const sheet = document.createElement("section");
      sheet.id = SHEET_ID;
      sheet.className = "verde-bologna-parks-sheet";
      sheet.setAttribute("aria-hidden", "true");
      document.body.appendChild(sheet);
    }
    return true;
  }

  function renderQuartiereFilters() {
    const node = $(FILTERS_ID)?.querySelector(".verde-bologna-parks-chips");
    if (!node) return;
    const counts = new Map();
    state.parks.forEach((record) => {
      const name = record.__vbQuartiere || "";
      if (name) counts.set(name, (counts.get(name) || 0) + 1);
    });
    const names = OFFICIAL_QUARTIERI.filter((name) => counts.has(name));
    node.innerHTML = ["", ...names].map((name) => {
      const label = name || "TUTTI";
      const count = name ? ` · ${counts.get(name) || 0}` : ` · ${state.parks.length}`;
      return `<button type="button" class="verde-bologna-parks-chip${state.activeQuartiere === name ? " is-active" : ""}" data-vb-quarter="${esc(name)}">${esc(label + count)}</button>`;
    }).join("");
    node.querySelectorAll("[data-vb-quarter]").forEach((button) => button.addEventListener("click", () => {
      state.activeQuartiere = button.getAttribute("data-vb-quarter") || "";
      renderQuartiereFilters();
      applyLiveFilter();
    }));
  }

  function formatFieldValue(value) {
    if (value === null || value === undefined || value === "") return "—";
    if (typeof value === "boolean") return value ? "Sì" : "No";
    if (typeof value === "object") {
      try { return JSON.stringify(value, null, 2); } catch (_) { return String(value); }
    }
    return String(value);
  }

  function isTechnicalGeometryField(key) {
    const normalized = normalizeText(key).replace(/[^a-z0-9]/g, "");
    return normalized.includes("geoshape") || normalized === "geometry" || normalized === "geom" || normalized.includes("geopoint");
  }

  function hasAreaBoundary(record) {
    return Boolean(areaGeometryOf(record));
  }

  function openCoboWorkOrder(record, triggerButton) {
    const workflow = window.HeraCoboMowing;
    if (typeof workflow?.openCreate !== "function") {
      window.alert("Sfalcio COBO non è disponibile in questo momento. Riprova tra qualche secondo.");
      return;
    }
    workflow.openCreate({
      record,
      parkName: parkName(record),
      parkCode: parkCodvia(record) || parkFallbackCode(record),
      quarter: record.__vbQuartiere || "",
      address: fieldValue(record, ["ubicazione", "indirizzo", "via", "nomevia", "nome_via"]),
      point: centerOf(record),
      boundaryAvailable: hasAreaBoundary(record),
      triggerButton
    });
  }

  function renderDetailSheet(record) {
    const sheet = $(SHEET_ID);
    if (!sheet) return;
    const codvia = parkCodvia(record) || "—";
    const name = parkName(record);
    const center = centerOf(record);
    const distance = recordDistance(record);
    const quarter = record.__vbQuartiere || "Non determinato";
    const boundaryAvailable = hasAreaBoundary(record);
    const boundaryStatus = boundaryAvailable ? "available" : (record.__vbBoundaryStatus || "unknown");
    const boundaryText = {
      available: "Disponibili sulla mappa",
      loading: "Cerco nel catasto comunale…",
      unavailable: "Non trovati nel catasto comunale",
      error: "Temporaneamente non disponibili",
      unknown: "Da verificare nel catasto comunale"
    }[boundaryStatus];
    const mapButtonText = boundaryAvailable ? "MOSTRA CONFINI" : (boundaryStatus === "loading" ? "CARICO CONFINI…" : "MOSTRA MAPPA");
    const mapButtonDisabled = !center || boundaryStatus === "loading";
    const fields = Object.entries(record).filter(([key]) => !String(key).startsWith("__vb") && !isTechnicalGeometryField(key));
    const navHref = center ? `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${center.lat},${center.lon}`)}` : "";
    sheet.innerHTML = `
      <header class="verde-bologna-parks-sheet-head">
        <button class="btn" type="button" data-vb-sheet-close>← INDIETRO</button>
        <h2>${esc(name)}</h2>
      </header>
      <section class="verde-bologna-parks-sheet-summary">
        <strong>CODVIA</strong><span>${esc(codvia)}</span>
        <strong>NOMEVIA</strong><span>${esc(name)}</span>
        <strong>QUARTIERE</strong><span>${esc(quarter)}</span>
        <strong>CONFINI</strong><span>${boundaryText}</span>
        ${Number.isFinite(distance) ? `<strong>DISTANZA</strong><span>${distance < 1000 ? `${Math.round(distance)} m` : `${(distance / 1000).toFixed(1)} km`}</span>` : ""}
      </section>
      <div class="verde-bologna-parks-sheet-actions">
        <button class="btn" type="button" data-vb-sheet-map ${mapButtonDisabled ? "disabled" : ""}>${mapButtonText}</button>
        ${center ? `<a class="btn btn-primary" href="${esc(navHref)}" target="_blank" rel="noopener">NAVIGA</a>` : `<button class="btn btn-primary" type="button" disabled>NAVIGA</button>`}
        <button class="btn" type="button" data-vb-create-cobo>🌿 CREA CANTIERE SFALCIO COBO</button>
      </div>
      <section class="verde-bologna-parks-fields">
        ${fields.map(([key, value]) => `<article class="verde-bologna-parks-field"><span>${esc(key)}</span>${typeof value === "object" && value !== null ? `<pre>${esc(formatFieldValue(value))}</pre>` : `<strong>${esc(formatFieldValue(value))}</strong>`}</article>`).join("")}
      </section>`;
    sheet.classList.add("is-open");
    sheet.setAttribute("aria-hidden", "false");
    sheet.querySelector("[data-vb-sheet-close]")?.addEventListener("click", closeDetailSheet);
    sheet.querySelector("[data-vb-sheet-map]")?.addEventListener("click", () => {
      closeDetailSheet();
      focusRecord(record, true);
    });
    sheet.querySelector("[data-vb-create-cobo]")?.addEventListener("click", (event) => openCoboWorkOrder(record, event.currentTarget));
  }

  function openDetailSheet(record) {
    const sheet = $(SHEET_ID);
    if (!sheet) return;
    state.listOpenRecord = record;
    if (areaGeometryOf(record)) record.__vbBoundaryStatus = "available";
    renderDetailSheet(record);
    if (!record.__vbBoundaryStatus || record.__vbBoundaryStatus === "unknown" || record.__vbBoundaryStatus === "error") {
      void loadManagedBoundary(record);
    }
  }

  function closeDetailSheet() {
    const sheet = $(SHEET_ID);
    sheet?.classList.remove("is-open");
    sheet?.setAttribute("aria-hidden", "true");
    state.listOpenRecord = null;
  }

  function renderList() {
    const node = $(LIST_ID);
    if (!node) return;
    if (state.loading) {
      node.innerHTML = `<p class="verde-bologna-parks-list-status">Carico i primi parchi e i filtri per quartiere…</p>`;
      return;
    }
    if (!state.filtered.length) {
      node.innerHTML = `<p class="verde-bologna-parks-list-status">Nessun risultato. Continua a scrivere CODVIA o NOMEVIA oppure cambia quartiere.</p>`;
      return;
    }
    const visibleRows = state.filtered.slice(0, LIST_RENDER_LIMIT);
    const locationText = state.userPosition ? "Ordinati dal più vicino al più lontano." : "Premi LA MIA POSIZIONE per ordinare dal più vicino.";
    const limitText = state.filtered.length > visibleRows.length ? ` Mostro i primi ${visibleRows.length}: usa la ricerca o il quartiere per restringere.` : "";
    node.innerHTML = `<p class="verde-bologna-parks-list-status">${state.filtered.length} risultati · ${esc(locationText + limitText)}</p>` + visibleRows.map((record, index) => `
      <button type="button" class="verde-bologna-parks-row" data-vb-park-index="${index}">
        <span class="verde-bologna-parks-code">${esc(parkCodvia(record) || "—")}</span>
        <span class="verde-bologna-parks-name">${esc(parkName(record))}</span>
        <span class="verde-bologna-parks-open" aria-hidden="true">›</span>
      </button>`).join("");
    node.querySelectorAll("[data-vb-park-index]").forEach((button) => button.addEventListener("click", () => {
      const record = state.filtered[Number(button.getAttribute("data-vb-park-index"))];
      if (record) openDetailSheet(record);
    }));
  }

  function markerCode(record) {
    const codvia = parkCodvia(record);
    if (codvia && !/^0(?:[.,]0+)?$/.test(codvia)) return { value: codvia, fallback: false };
    const fallback = parkFallbackCode(record);
    return fallback ? { value: fallback, fallback: true } : { value: codvia || "—", fallback: true };
  }

  function ensureLayer() {
    if (!state.map || !window.L) return;
    if (!state.boundaryLayer) state.boundaryLayer = window.L.layerGroup().addTo(state.map);
    if (!state.layer) state.layer = window.L.layerGroup().addTo(state.map);
    if (!state.markerRenderer && window.L.canvas) state.markerRenderer = window.L.canvas({ padding: 0.5 });
    if (!state.mapEventsBound) {
      const scheduleViewportRefresh = () => {
        if (!parksActive() || !state.loaded) return;
        window.clearTimeout(state.markerRefreshTimer);
        state.markerRefreshTimer = window.setTimeout(() => renderMapMarkers({ fitMap: false }), 120);
      };
      state.map.on("zoomend", scheduleViewportRefresh);
      state.map.on("moveend", () => {
        if (state.map?.getZoom?.() >= 15) scheduleViewportRefresh();
      });
      state.mapEventsBound = true;
    }
  }

  function renderMapMarkers({ fitMap = true } = {}) {
    if (!parksActive()) return;
    ensureLayer();
    if (!state.layer || !state.map || !window.L) return;
    state.layer.clearLayers();
    const bounds = window.L.latLngBounds([]);
    const viewport = !fitMap && state.map.getZoom() >= 15 ? state.map.getBounds()?.pad?.(0.15) : null;
    const visibleRecords = viewport?.isValid?.()
      ? state.filtered.filter((record) => {
        const center = centerOf(record);
        return center ? viewport.contains([center.lat, center.lon]) : false;
      })
      : state.filtered;
    const showCodeLabels = visibleRecords.length <= LABEL_MARKER_LIMIT;
    visibleRecords.forEach((record) => {
      const center = centerOf(record);
      if (!center) return;
      const code = markerCode(record);
      const marker = showCodeLabels
        ? window.L.marker([center.lat, center.lon], {
          icon: window.L.divIcon({
            className: "vb-park-live-wrap",
            html: `<span class="vb-park-live-code${code.fallback ? " is-fallback" : ""}">${esc(code.value)}</span>`,
            iconSize: null,
            iconAnchor: [21, 14]
          }),
          keyboard: true,
          riseOnHover: true,
          title: `${code.value} · ${parkName(record)}`
        })
        : window.L.circleMarker([center.lat, center.lon], {
          renderer: state.markerRenderer,
          radius: 6,
          color: "#08783f",
          weight: 2,
          fillColor: "#31b96b",
          fillOpacity: 0.78
        });
      marker.addTo(state.layer);
      marker.bindPopup(`<strong>CODVIA ${esc(parkCodvia(record) || "—")}</strong><br>${esc(parkName(record))}`);
      marker.on("click", () => openDetailSheet(record));
      bounds.extend([center.lat, center.lon]);
    });
    if (fitMap && bounds.isValid()) {
      const query = String($("verde-bologna-query")?.value || "").trim();
      const maxZoom = query || state.activeQuartiere ? 17 : 14;
      state.map.fitBounds(bounds.pad(0.08), { animate: false, maxZoom });
    }
  }

  function focusRecord(record, openPopup) {
    const center = centerOf(record);
    if (!center || !state.map) return;
    ensureLayer();
    state.boundaryLayer?.clearLayers?.();
    const geometry = areaGeometryOf(record);
    let boundaryBounds = null;
    if (hasAreaBoundary(record) && state.boundaryLayer && window.L) {
      const boundary = window.L.geoJSON({ type: "Feature", geometry, properties: {} }, {
        style: { color: "#0b6b3a", weight: 4, fillColor: "#45a96a", fillOpacity: 0.2 }
      }).addTo(state.boundaryLayer);
      boundaryBounds = boundary.getBounds?.();
    }
    if (boundaryBounds?.isValid?.()) state.map.fitBounds(boundaryBounds.pad(0.08), { animate: false, maxZoom: 18 });
    else state.map.setView([center.lat, center.lon], Math.max(state.map.getZoom(), 17), { animate: false });
    if (openPopup) {
      const boundaryText = boundaryBounds?.isValid?.() ? "<br>Confine ufficiale evidenziato" : "";
      const popup = window.L.popup().setLatLng([center.lat, center.lon]).setContent(`<strong>CODVIA ${esc(parkCodvia(record) || "—")}</strong><br>${esc(parkName(record))}${boundaryText}`).openOn(state.map);
      void popup;
    }
    $("verde-bologna-map-card")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function applyLiveFilter() {
    if (!parksActive() || !state.loaded) return;
    const query = normalizeText($("verde-bologna-query")?.value || "");
    const compactQuery = query.replace(/\s+/g, "");
    const positionKey = state.userPosition ? `${state.userPosition.lat.toFixed(5)},${state.userPosition.lon.toFixed(5)}` : "";
    const signature = `${query}|${state.activeQuartiere}|${positionKey}|${state.parks.length}`;
    if (signature === state.renderSignature) return;
    let rows = state.parks.filter((record) => {
      if (state.activeQuartiere && record.__vbQuartiere !== state.activeQuartiere) return false;
      if (!query) return true;
      const codvia = normalizeText(record.__vbCodvia).replace(/\s+/g, "");
      const name = normalizeText(record.__vbName);
      const codeMatch = compactQuery && codvia.startsWith(compactQuery);
      const nameMatch = record.__vbSearch.includes(query) || name.includes(query);
      return codeMatch || nameMatch;
    });
    rows = rows.sort((a, b) => {
      if (state.userPosition) {
        const distanceDiff = recordDistance(a) - recordDistance(b);
        if (Math.abs(distanceDiff) > 0.01) return distanceDiff;
      }
      const codeA = parkCodvia(a);
      const codeB = parkCodvia(b);
      if (/^\d+$/.test(codeA) && /^\d+$/.test(codeB)) return Number(codeA) - Number(codeB);
      return parkName(a).localeCompare(parkName(b), "it", { sensitivity: "base", numeric: true });
    });
    state.filtered = rows;
    state.renderSignature = signature;
    state.boundaryLayer?.clearLayers?.();
    renderList();
    renderMapMarkers();
    const status = $("verde-bologna-status");
    if (status) status.textContent = `${rows.length} parchi/giardini corrispondenti ai filtri. Ricerca live per CODVIA e NOMEVIA.`;
  }

  async function loadParksData() {
    if (state.loaded) {
      state.renderSignature = "";
      applyLiveFilter();
      return;
    }
    if (state.loading) return;
    const generation = state.generation;
    state.loading = true;
    renderList();
    try {
      const cacheKey = `varga-verde-bologna:all:${PARKS_DATASET_ID}`;
      const cached = readSessionCache(cacheKey);
      if (Array.isArray(cached) && cached.length) {
        if (generation !== state.generation || !parksActive()) return;
        publishParkRecords(cached, { complete: true });
        return;
      }

      const firstPage = await waitForBasePageCache() || await fetchRecordsPage(PARKS_DATASET_ID, 0);
      if (generation !== state.generation || !parksActive()) return;
      publishParkRecords(firstPage.records);

      const remainingOffsets = [];
      for (let offset = firstPage.records.length; offset < firstPage.total; offset += PAGE_SIZE) remainingOffsets.push(offset);
      if (!remainingOffsets.length) {
        state.complete = true;
        writeSessionCache(cacheKey, firstPage.records);
        return;
      }

      const status = $("verde-bologna-status");
      if (status && parksActive()) status.textContent = `${firstPage.records.length} parchi disponibili. Completo l’elenco in background…`;
      const settledPages = await Promise.allSettled(remainingOffsets.map((offset) => fetchRecordsPage(PARKS_DATASET_ID, offset)));
      if (generation !== state.generation || !parksActive()) return;
      const complete = settledPages.every((result) => result.status === "fulfilled");
      const records = [
        ...firstPage.records,
        ...settledPages.flatMap((result) => result.status === "fulfilled" ? result.value.records : [])
      ];
      publishParkRecords(records, { complete });
      if (complete) {
        writeSessionCache(cacheKey, records);
      } else if (status && parksActive()) {
        status.textContent = `${records.length} parchi disponibili. Alcuni dati non hanno risposto: puoi comunque usare ricerca, quartieri e mappa.`;
      }
    } catch (error) {
      state.loading = false;
      if (generation !== state.generation || !parksActive()) return;
      const node = $(LIST_ID);
      if (node) node.innerHTML = `<p class="verde-bologna-parks-list-status">${esc(error?.message || "Impossibile caricare l'elenco completo dei parchi.")}</p>`;
      const status = $("verde-bologna-status");
      if (status) status.textContent = error?.message || "Impossibile caricare i parchi in questo momento.";
    }
  }

  function requestUserPosition() {
    if (!parksActive() || state.locationRequested || !navigator.geolocation) return;
    state.locationRequested = true;
    navigator.geolocation.getCurrentPosition((position) => {
      const lat = Number(position.coords.latitude);
      const lon = Number(position.coords.longitude);
      if (Number.isFinite(lat) && Number.isFinite(lon)) {
        state.userPosition = { lat, lon };
        applyLiveFilter();
      }
    }, () => {
      state.locationRequested = false;
      renderList();
    }, { enableHighAccuracy: true, timeout: 10000, maximumAge: 60000 });
  }

  function activateParksMode() {
    const page = $(PAGE_ID);
    if (!page || !parksActive()) return;
    if (state.active && page.classList.contains("vb-parks-advanced")) return;
    state.active = true;
    ensureUi();
    page.classList.add("vb-parks-advanced");
    const input = $("verde-bologna-query");
    if (input) {
      input.placeholder = "Cerca NOMEVIA o CODVIA…";
      input.setAttribute("inputmode", "search");
      input.setAttribute("autocomplete", "off");
    }
    renderQuartiereFilters();
    void loadParksData();
  }

  function deactivateParksMode() {
    const page = $(PAGE_ID);
    const hasAdvancedClass = page?.classList.contains("vb-parks-advanced");
    if (!state.active && !hasAdvancedClass) return;
    if (hasAdvancedClass) page.classList.remove("vb-parks-advanced");
    if (!state.active) return;
    state.active = false;
    state.generation += 1;
    state.abortControllers.forEach((controller) => controller.abort());
    state.abortControllers.clear();
    state.loading = false;
    if (!state.complete) state.loaded = false;
    window.clearTimeout(state.filterTimer);
    window.clearTimeout(state.markerRefreshTimer);
    state.layer?.clearLayers?.();
    state.boundaryLayer?.clearLayers?.();
    closeDetailSheet();
  }

  function resetCategoryState(event) {
    if (event?.detail?.map && state.map && event.detail.map !== state.map) return;
    state.generation += 1;
    state.abortControllers.forEach((controller) => controller.abort());
    state.abortControllers.clear();
    window.clearTimeout(state.markerRefreshTimer);
    window.clearTimeout(state.filterTimer);
    state.layer?.clearLayers?.();
    state.boundaryLayer?.clearLayers?.();
    state.map = null;
    state.layer = null;
    state.boundaryLayer = null;
    state.markerRenderer = null;
    state.mapEventsBound = false;
    state.parks = [];
    state.quartieri = [];
    state.filtered = [];
    state.loading = false;
    state.loaded = false;
    state.complete = false;
    state.activeQuartiere = "";
    state.userPosition = null;
    state.locationRequested = false;
    state.renderSignature = "";
    closeDetailSheet();
    state.active = false;
    deactivateParksMode();
  }

  function captureCreatedMap(event) {
    const map = event?.detail?.map;
    if (!map) return;
    state.map = map;
    ensureLayer();
    if (parksActive()) window.setTimeout(() => renderMapMarkers(), 80);
  }

  function handleCategoryOpened() {
    window.setTimeout(() => {
      if (parksActive()) activateParksMode(); else deactivateParksMode();
    }, 0);
  }

  function installSearchBehavior() {
    const input = $("verde-bologna-query");
    const form = $("verde-bologna-search-form");
    const clear = $("verde-bologna-clear-btn");
    if (!input || !form || input.dataset.vbParksLive === "1") return;
    input.dataset.vbParksLive = "1";
    input.addEventListener("input", () => {
      if (!parksActive()) return;
      window.clearTimeout(state.filterTimer);
      state.filterTimer = window.setTimeout(applyLiveFilter, SEARCH_DEBOUNCE_MS);
    });
    form.addEventListener("submit", (event) => {
      if (!parksActive()) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      window.clearTimeout(state.filterTimer);
      applyLiveFilter();
    }, true);
    clear?.addEventListener("click", (event) => {
      if (!parksActive()) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      input.value = "";
      state.activeQuartiere = "";
      state.renderSignature = "";
      renderQuartiereFilters();
      applyLiveFilter();
      input.focus();
    }, true);
    $("verde-bologna-location-btn")?.addEventListener("click", () => {
      if (!parksActive() || !navigator.geolocation) return;
      state.locationRequested = false;
      window.setTimeout(requestUserPosition, 50);
    });
  }

  function installCategoryBehavior() {
    const page = $(PAGE_ID);
    const select = $(CATEGORY_ID);
    if (!page || !select || select.dataset.vbParksCategory === "1") return;
    select.dataset.vbParksCategory = "1";
    select.addEventListener("change", () => window.setTimeout(() => {
      if (parksActive()) activateParksMode(); else deactivateParksMode();
    }, 80));
    const observer = new MutationObserver(() => {
      if (parksActive()) activateParksMode(); else if (page.classList.contains("hidden")) deactivateParksMode();
    });
    observer.observe(page, { attributes: true, attributeFilter: ["class", "aria-hidden"] });
  }

  function install() {
    if (state.installDone) return;
    removeLegacyQuartieriCache();
    injectStyle();
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      if (ensureUi() && $(CATEGORY_ID)) {
        installSearchBehavior();
        installCategoryBehavior();
        state.installDone = true;
        window.clearInterval(timer);
        if (parksActive()) activateParksMode();
      } else if (attempts > 120) {
        window.clearInterval(timer);
      }
    }, 100);
  }

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && $(SHEET_ID)?.classList.contains("is-open")) {
      event.stopPropagation();
      closeDetailSheet();
    }
  }, true);

  window.addEventListener("hera:verde-bologna-map-destroyed", resetCategoryState);
  window.addEventListener("hera:verde-bologna-map-created", captureCreatedMap);
  window.addEventListener("hera:verde-bologna-category-opened", handleCategoryOpened);
  window.addEventListener("hera:verde-bologna-category-closed", deactivateParksMode);

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();

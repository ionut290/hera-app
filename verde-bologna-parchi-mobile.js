(() => {
  "use strict";

  const PAGE_ID = "verde-bologna-page";
  const CATEGORY_ID = "verde-bologna-operativo-category";
  const PARKS_DATASET_ID = "carta-tecnica-comunale-toponimi-parchi-e-giardini";
  const QUARTIERI_DATASET_ID = "quartieri-di-bologna";
  const API_ROOT = "https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets";
  const MOBILE_QUERY = "(max-width: 760px)";
  const STYLE_ID = "verde-bologna-parchi-mobile-style";
  const FILTERS_ID = "verde-bologna-parchi-quartieri";
  const LIST_ID = "verde-bologna-parchi-list";
  const SHEET_ID = "verde-bologna-parchi-sheet";
  const CACHE_TTL_MS = 10 * 60 * 1000;
  const SEARCH_DEBOUNCE_MS = 180;
  const LIST_RENDER_LIMIT = 60;
  const LABEL_MARKER_LIMIT = 80;
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
    parks: [],
    quartieri: [],
    filtered: [],
    loading: false,
    loaded: false,
    activeQuartiere: "",
    userPosition: null,
    locationRequested: false,
    listOpenRecord: null,
    mapFactoryWrapped: false,
    installDone: false,
    markerRenderer: null,
    mapEventsBound: false,
    markerRefreshTimer: 0,
    filterTimer: 0,
    renderSignature: ""
  };

  const $ = (id) => document.getElementById(id);
  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[char]));

  function mobileActive() {
    return window.matchMedia(MOBILE_QUERY).matches;
  }

  function parksActive() {
    return mobileActive() && $(CATEGORY_ID)?.value === PARKS_DATASET_ID && !$(PAGE_ID)?.classList.contains("hidden");
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

  function geometryOf(record) {
    for (const [key, value] of Object.entries(record || {})) {
      const normalized = normalizeText(key).replace(/[^a-z0-9]/g, "");
      if (!normalized.includes("geoshape") && normalized !== "geometry" && normalized !== "geom" && !normalized.includes("geopoint")) continue;
      const geometry = parseGeometryValue(value);
      if (geometry) return geometry;
    }
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

  async function fetchAllRecords(datasetId) {
    const cacheKey = `varga-verde-bologna:all:${datasetId}`;
    const cached = readSessionCache(cacheKey);
    if (Array.isArray(cached)) return cached;
    const records = [];
    let offset = 0;
    let total = Number.POSITIVE_INFINITY;
    while (offset < total && offset < 10000) {
      const params = new URLSearchParams({ limit: "100", offset: String(offset) });
      const response = await fetch(`${API_ROOT}/${encodeURIComponent(datasetId)}/records?${params}`, { headers: { Accept: "application/json" } });
      if (!response.ok) throw new Error(`Open Data Comune di Bologna non disponibili (${response.status}).`);
      const payload = await response.json();
      const page = Array.isArray(payload?.results) ? payload.results : [];
      total = Number(payload?.total_count ?? page.length);
      records.push(...page);
      if (!page.length) break;
      offset += page.length;
    }
    writeSessionCache(cacheKey, records);
    return records;
  }

  function injectStyle() {
    if ($(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
      .verde-bologna-parks-tools,.verde-bologna-parks-list,.verde-bologna-parks-sheet{display:none}
      @media ${MOBILE_QUERY}{
        .verde-bologna-page.vb-parks-advanced .verde-bologna-code-marker-wrap{display:none!important}
        .verde-bologna-page.vb-parks-advanced #verde-bologna-results,
        .verde-bologna-page.vb-parks-advanced #verde-bologna-load-more{display:none!important}
        .verde-bologna-page.vb-parks-advanced .verde-bologna-search{margin-bottom:6px!important}
        .verde-bologna-parks-tools{display:block;margin:0 0 10px}
        .verde-bologna-parks-tools-title{margin:0 0 6px;font-size:.72rem;font-weight:900;color:#47617d;text-transform:uppercase;letter-spacing:.04em}
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

  function openDetailSheet(record) {
    const sheet = $(SHEET_ID);
    if (!sheet) return;
    state.listOpenRecord = record;
    const codvia = parkCodvia(record) || "—";
    const name = parkName(record);
    const center = centerOf(record);
    const distance = recordDistance(record);
    const quarter = record.__vbQuartiere || "Non determinato";
    const fields = Object.entries(record).filter(([key]) => !String(key).startsWith("__vb"));
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
        ${Number.isFinite(distance) ? `<strong>DISTANZA</strong><span>${distance < 1000 ? `${Math.round(distance)} m` : `${(distance / 1000).toFixed(1)} km`}</span>` : ""}
      </section>
      <div class="verde-bologna-parks-sheet-actions">
        <button class="btn" type="button" data-vb-sheet-map ${center ? "" : "disabled"}>MOSTRA MAPPA</button>
        ${center ? `<a class="btn btn-primary" href="${esc(navHref)}" target="_blank" rel="noopener">NAVIGA</a>` : `<button class="btn btn-primary" type="button" disabled>NAVIGA</button>`}
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
      node.innerHTML = `<p class="verde-bologna-parks-list-status">Carico tutti i parchi e i quartieri ufficiali del Comune di Bologna…</p>`;
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
    state.map.setView([center.lat, center.lon], Math.max(state.map.getZoom(), 17), { animate: false });
    if (openPopup) {
      const popup = window.L.popup().setLatLng([center.lat, center.lon]).setContent(`<strong>CODVIA ${esc(parkCodvia(record) || "—")}</strong><br>${esc(parkName(record))}`).openOn(state.map);
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
    state.loading = true;
    renderList();
    try {
      const parks = await fetchAllRecords(PARKS_DATASET_ID);
      state.parks = parks.map(prepareRecord);
      state.loaded = true;
      state.loading = false;
      renderQuartiereFilters();
      applyLiveFilter();
    } catch (error) {
      state.loading = false;
      const node = $(LIST_ID);
      if (node) node.innerHTML = `<p class="verde-bologna-parks-list-status">${esc(error?.message || "Impossibile caricare l'elenco completo dei parchi.")}</p>`;
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
    $(PAGE_ID)?.classList.remove("vb-parks-advanced");
    window.clearTimeout(state.filterTimer);
    state.layer?.clearLayers?.();
    closeDetailSheet();
  }

  function captureMapFactory() {
    if (state.mapFactoryWrapped || !window.L?.map) return Boolean(state.mapFactoryWrapped);
    const previousMapFactory = window.L.map;
    if (previousMapFactory.__vbParksCapture) {
      state.mapFactoryWrapped = true;
      return true;
    }
    const wrapped = function(element, options) {
      const map = previousMapFactory(element, options);
      const node = typeof element === "string" ? document.getElementById(element) : element;
      if (node?.id === "verde-bologna-map") {
        state.map = map;
        ensureLayer();
        if (parksActive()) window.setTimeout(() => renderMapMarkers(), 80);
      }
      return map;
    };
    wrapped.__vbParksCapture = true;
    wrapped.__previousMapFactory = previousMapFactory;
    window.L.map = wrapped;
    state.mapFactoryWrapped = true;
    return true;
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
    page.addEventListener("click", (event) => {
      const button = event.target?.closest?.("[data-vb-open]");
      if (!button) return;
      window.setTimeout(() => {
        if (parksActive()) activateParksMode(); else deactivateParksMode();
      }, 100);
    }, true);
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
      captureMapFactory();
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

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();

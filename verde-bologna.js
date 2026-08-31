(() => {
  "use strict";

  const API_ROOT = "https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets";
  const PAGE_SIZE = 100;
  const CACHE_TTL_MS = 10 * 60 * 1000;
  const PAGE_ID = "verde-bologna-page";
  const STYLE_ID = "verde-bologna-style";
  const MENU_BUTTON_ID = "open-verde-bologna-btn";
  const CACHE_PREFIX = "varga-verde-bologna:";

  const DATASETS = Object.freeze([
    { id: "un_gest", icon: "🌳", title: "Aree verdi in manutenzione", short: "Aiuole, parchi, giardini, verde scolastico, sportivo e stradale.", priority: true },
    { id: "alberi-manutenzioni", icon: "🌲", title: "Alberi singoli", short: "Catasto degli alberi con numero punto, specie, caratteristiche e coordinate.", delegate: "open-tree-search-btn" },
    { id: "popolazione-arborea", icon: "🌴", title: "Popolazioni arboree", short: "Gruppi e superfici con popolazioni arboree e arbustive." },
    { id: "siepi", icon: "🌿", title: "Siepi in manutenzione", short: "Specie, tipologia, lunghezza, altezza, larghezza e superficie di potatura." },
    { id: "attrezzature_ludiche_ginniche_sportive", icon: "🛝", title: "Giochi e attrezzature sportive", short: "Attrezzature ludiche, ginniche e sportive presenti sul territorio." },
    { id: "arredo", icon: "🪑", title: "Arredo urbano comunale", short: "Arredi censiti dal Comune di Bologna, separati dai dati OSM." },
    { id: "sgambatura_cani", icon: "🐕", title: "Aree cani", short: "Aree di sgambatura cani in manutenzione comunale." },
    { id: "carta-tecnica-comunale-toponimi-parchi-e-giardini", icon: "🏞️", title: "Parchi e giardini", short: "Toponimi ufficiali dei parchi e giardini del Comune." },
    { id: "aree-verdi_entrate_centroidi", icon: "🚪", title: "Ingressi aree verdi", short: "Centroidi e ingressi delle maggiori aree verdi, utili per la navigazione." },
    { id: "aree-ortive", icon: "🥕", title: "Aree ortive", short: "Orti comunali, gestori, indirizzi e informazioni disponibili." },
    { id: "verde_privato_urbanizzato", icon: "🏡", title: "Verde privato", short: "Verde privato nel territorio urbanizzato, mantenuto separato dal verde pubblico.", privateGreen: true }
  ]);

  const state = {
    datasetId: "un_gest", query: "", offset: 0, total: 0, records: [], map: null,
    baseLayer: null, featureLayer: null, featureByIndex: new Map(), userMarker: null,
    userAccuracy: null, requestSerial: 0, fullscreen: false
  };

  const $ = (id) => document.getElementById(id);
  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[char]));

  function sourcePageUrl(datasetId) {
    return `https://opendata.comune.bologna.it/explore/dataset/${encodeURIComponent(datasetId)}/`;
  }

  function apiUrl(datasetId, offset = 0, query = "", includeSearch = true) {
    const params = new URLSearchParams({ limit: String(PAGE_SIZE), offset: String(offset) });
    if (query && includeSearch) {
      const safe = String(query).replace(/\\/g, "\\\\").replace(/"/g, '\\"');
      params.set("where", `search(\"${safe}\")`);
    }
    return `${API_ROOT}/${encodeURIComponent(datasetId)}/records?${params.toString()}`;
  }

  function cacheKey(datasetId, offset, query, includeSearch) {
    return `${CACHE_PREFIX}${datasetId}:${offset}:${includeSearch ? "server" : "plain"}:${String(query).toLocaleLowerCase("it-IT")}`;
  }

  function readCache(datasetId, offset, query, includeSearch) {
    try {
      const raw = sessionStorage.getItem(cacheKey(datasetId, offset, query, includeSearch));
      if (!raw) return null;
      const cached = JSON.parse(raw);
      if (!cached || Date.now() - Number(cached.savedAt || 0) > CACHE_TTL_MS) return null;
      return cached.payload || null;
    } catch (_) { return null; }
  }

  function writeCache(datasetId, offset, query, includeSearch, payload) {
    try { sessionStorage.setItem(cacheKey(datasetId, offset, query, includeSearch), JSON.stringify({ savedAt: Date.now(), payload })); } catch (_) {}
  }

  function injectStyle() {
    if ($(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
      .verde-bologna-page{position:fixed;inset:0;z-index:1060;overflow:auto;background:#eef5f0;color:#173426}
      .verde-bologna-page.hidden{display:none!important}.verde-bologna-shell{width:min(1180px,100%);margin:auto;padding:0 16px 28px}
      .verde-bologna-header{position:sticky;top:0;z-index:20;display:grid;grid-template-columns:auto minmax(0,1fr) auto;gap:14px;align-items:center;margin:0 -16px 18px;padding:max(12px,env(safe-area-inset-top)) max(16px,calc((100vw - 1180px)/2 + 16px)) 12px;background:rgba(255,255,255,.96);border-bottom:1px solid #cdded2;backdrop-filter:blur(14px)}
      .verde-bologna-header h1{margin:0;font-size:clamp(1.25rem,3vw,1.8rem);color:#154d2e}.verde-bologna-header p{margin:3px 0 0;color:#5f7868;font-size:.88rem}.verde-bologna-badge{padding:7px 10px;border-radius:999px;background:#e0f4e6;color:#146435;font-size:.76rem;font-weight:900;white-space:nowrap}
      .verde-bologna-hero{display:grid;grid-template-columns:minmax(0,1.25fr) minmax(280px,.75fr);gap:16px;margin-bottom:18px}.verde-bologna-hero-card,.verde-bologna-note{padding:18px;border:1px solid #cfe0d3;border-radius:20px;background:#fff;box-shadow:0 8px 24px rgba(31,78,47,.08)}
      .verde-bologna-hero-card h2,.verde-bologna-note h2{margin:0 0 7px;color:#174d30}.verde-bologna-hero-card p,.verde-bologna-note p{margin:0;color:#526d5b;line-height:1.5}.verde-bologna-note{background:#f7fbf8}.verde-bologna-note strong{color:#8a4b0f}
      .verde-bologna-section-title{display:flex;align-items:end;justify-content:space-between;gap:12px;margin:22px 0 10px}.verde-bologna-section-title h2{margin:0;color:#173f29}.verde-bologna-section-title span{color:#698072;font-size:.83rem}.verde-bologna-datasets{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}
      .verde-bologna-dataset{display:grid;grid-template-rows:auto auto 1fr auto auto;gap:8px;min-height:190px;padding:15px;border:1px solid #ccdcd0;border-radius:18px;background:#fff;text-align:left;box-shadow:0 7px 18px rgba(29,80,47,.07)}.verde-bologna-dataset.is-active{border:2px solid #18854b;background:#f3fbf6}.verde-bologna-dataset.is-private{border-color:#e2c68f;background:#fffaf0}
      .verde-bologna-dataset-icon{font-size:1.65rem}.verde-bologna-dataset h3{margin:0;color:#1e4d31;font-size:1rem}.verde-bologna-dataset p{margin:0;color:#5e7464;line-height:1.42;font-size:.88rem}.verde-bologna-dataset small{color:#74877a;font-size:.72rem;overflow-wrap:anywhere}.verde-bologna-dataset .btn{margin-top:4px;width:100%;min-height:40px}
      .verde-bologna-browser{margin-top:18px;padding:16px;border:1px solid #cdded2;border-radius:20px;background:#fff;box-shadow:0 8px 24px rgba(31,78,47,.08)}.verde-bologna-browser-head{display:flex;justify-content:space-between;gap:14px;align-items:flex-start}.verde-bologna-browser-head h2{margin:0;color:#164d2e}.verde-bologna-browser-head p{margin:5px 0 0;color:#667d6d}
      .verde-bologna-source-link{display:inline-flex;align-items:center;justify-content:center;min-height:42px;padding:9px 12px;border:1px solid #9ec8aa;border-radius:12px;color:#155e35;background:#f1faf4;font-weight:900;text-decoration:none;white-space:nowrap}.verde-bologna-search{display:grid;grid-template-columns:minmax(0,1fr) auto auto;gap:8px;margin:14px 0}.verde-bologna-search input{min-height:46px;padding:10px 12px;border:1px solid #aebfb2;border-radius:11px;font:inherit}
      .verde-bologna-status{margin:0 0 12px;padding:9px 11px;border-radius:10px;background:#edf7f0;color:#315b3e;font-size:.85rem}.verde-bologna-status.error{background:#fff0ef;color:#9b281f}.verde-bologna-status.warning{background:#fff7df;color:#7b5304}.verde-bologna-map-card{display:grid;gap:8px;margin:12px 0 16px;padding:12px;border:1px solid #cdded2;border-radius:16px;background:#f9fcfa}
      .verde-bologna-map-toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.verde-bologna-map-toolbar strong{margin-right:auto}.verde-bologna-map{height:min(48vh,520px);min-height:330px;border-radius:12px;overflow:hidden}.verde-bologna-map-status{margin:0;color:#5d7464;font-size:.8rem}body.verde-bologna-fullscreen-open{overflow:hidden}.verde-bologna-map-card.is-fullscreen{position:fixed;inset:0;z-index:12060;margin:0;padding:max(8px,env(safe-area-inset-top)) max(8px,env(safe-area-inset-right)) max(8px,env(safe-area-inset-bottom)) max(8px,env(safe-area-inset-left));border-radius:0;background:#eef5f0;grid-template-rows:auto auto minmax(0,1fr)}.verde-bologna-map-card.is-fullscreen .verde-bologna-map{height:100%;min-height:0}
      .verde-bologna-results{display:grid;gap:10px}.verde-bologna-result{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:12px;padding:13px;border:1px solid #d6e3d9;border-radius:14px;background:#fbfdfb}.verde-bologna-result h3{margin:0;color:#174d30;font-size:1rem}.verde-bologna-result p{margin:4px 0 0;color:#617667;font-size:.84rem;line-height:1.4}.verde-bologna-result-actions{display:flex;gap:6px;align-items:center;flex-wrap:wrap;justify-content:flex-end}.verde-bologna-result-actions .btn,.verde-bologna-result-actions a{min-height:38px;padding:7px 10px;font-size:.76rem;text-decoration:none}
      .verde-bologna-details{grid-column:1/-1;display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:7px}.verde-bologna-details div{padding:8px;border-radius:9px;background:#eff6f1}.verde-bologna-details span{display:block;color:#6c7f70;font-size:.7rem}.verde-bologna-details strong{display:block;margin-top:3px;color:#294d35;font-size:.8rem;overflow-wrap:anywhere}.verde-bologna-load-more{display:block;width:100%;margin-top:12px;min-height:46px}.verde-bologna-empty{padding:18px;text-align:center;color:#65796a}
      @media(max-width:760px){.verde-bologna-header{grid-template-columns:auto minmax(0,1fr)}.verde-bologna-badge{grid-column:1/-1;justify-self:start}.verde-bologna-hero{grid-template-columns:1fr}.verde-bologna-search{grid-template-columns:1fr 1fr}.verde-bologna-search input{grid-column:1/-1}.verde-bologna-browser-head{display:grid}.verde-bologna-source-link{width:100%}.verde-bologna-result{grid-template-columns:1fr}.verde-bologna-result-actions{justify-content:flex-start}}
      @media(max-width:520px){.verde-bologna-shell{padding:0 10px 20px}.verde-bologna-header{margin:0 -10px 12px;padding-left:10px;padding-right:10px}.verde-bologna-header .btn{padding:7px 9px}.verde-bologna-datasets{grid-template-columns:1fr}.verde-bologna-search{grid-template-columns:1fr}.verde-bologna-search input{grid-column:auto}.verde-bologna-map{height:44vh;min-height:300px}}
    `;
    document.head.appendChild(style);
  }

  function buildPage() {
    if ($(PAGE_ID)) return $(PAGE_ID);
    const page = document.createElement("section");
    page.id = PAGE_ID;
    page.className = "verde-bologna-page hidden";
    page.setAttribute("aria-hidden", "true");
    page.innerHTML = `
      <div class="verde-bologna-shell">
        <header class="verde-bologna-header">
          <button id="verde-bologna-back-btn" class="btn" type="button">← HOME</button>
          <div><h1>🌳 Verde Bologna</h1><p>Catasto verde ufficiale del Comune di Bologna</p></div>
          <span class="verde-bologna-badge">COMUNE DI BOLOGNA OPEN DATA</span>
        </header>
        <section class="verde-bologna-hero">
          <div class="verde-bologna-hero-card"><h2>Un solo punto per il verde comunale</h2><p>Aree verdi, alberi, siepi, arredi, giochi, aree cani, parchi, ingressi, orti e verde privato vengono interrogati direttamente dai dataset pubblici del Comune di Bologna. I dati sono caricati solo quando apri una categoria.</p></div>
          <div class="verde-bologna-note"><h2>Priorità ai dati ufficiali</h2><p><strong>Per Bologna il Comune è la fonte principale.</strong> OpenStreetMap resta disponibile nelle altre sezioni dell’app come integrazione cartografica, ma qui non sostituisce il censimento comunale.</p></div>
        </section>
        <div class="verde-bologna-section-title"><h2>Dataset disponibili</h2><span>11 categorie ufficiali</span></div>
        <section id="verde-bologna-datasets" class="verde-bologna-datasets" aria-label="Dataset del verde di Bologna"></section>
        <section id="verde-bologna-browser" class="verde-bologna-browser hidden" aria-live="polite">
          <div class="verde-bologna-browser-head">
            <div><h2 id="verde-bologna-active-title">Dataset</h2><p id="verde-bologna-active-description"></p></div>
            <a id="verde-bologna-source-link" class="verde-bologna-source-link" href="#" target="_blank" rel="noopener noreferrer">FONTE UFFICIALE ↗</a>
          </div>
          <form id="verde-bologna-search-form" class="verde-bologna-search">
            <input id="verde-bologna-query" type="search" autocomplete="off" placeholder="Cerca nel dataset selezionato...">
            <button class="btn btn-primary" type="submit">CERCA</button>
            <button id="verde-bologna-clear-btn" class="btn" type="button">AZZERA</button>
          </form>
          <p id="verde-bologna-status" class="verde-bologna-status" role="status">Seleziona un dataset.</p>
          <section id="verde-bologna-map-card" class="verde-bologna-map-card" aria-label="Mappa Verde Bologna">
            <div class="verde-bologna-map-toolbar"><strong>Mappa dei risultati caricati</strong><button id="verde-bologna-location-btn" class="btn" type="button">⌖ LA MIA POSIZIONE</button><button id="verde-bologna-fullscreen-btn" class="btn" type="button" aria-pressed="false">⛶ SCHERMO INTERO</button></div>
            <p id="verde-bologna-map-status" class="verde-bologna-map-status">La mappa usa OpenStreetMap come sfondo; i dati del verde provengono dal Comune di Bologna.</p>
            <div id="verde-bologna-map" class="verde-bologna-map"></div>
          </section>
          <section id="verde-bologna-results" class="verde-bologna-results"></section>
          <button id="verde-bologna-load-more" class="btn verde-bologna-load-more hidden" type="button">CARICA ALTRI 100</button>
        </section>
      </div>`;
    document.body.appendChild(page);
    return page;
  }

  function renderDatasetCards() {
    const node = $("verde-bologna-datasets");
    if (!node) return;
    node.innerHTML = DATASETS.map((dataset) => `
      <article class="verde-bologna-dataset${dataset.id === state.datasetId ? " is-active" : ""}${dataset.privateGreen ? " is-private" : ""}" data-vb-dataset-card="${esc(dataset.id)}">
        <span class="verde-bologna-dataset-icon" aria-hidden="true">${dataset.icon}</span>
        <h3>${esc(dataset.title)}${dataset.priority ? " · PRIORITÀ" : ""}</h3>
        <p>${esc(dataset.short)}</p><small>${esc(dataset.id)}</small>
        <button class="btn${dataset.priority ? " btn-primary" : ""}" type="button" data-vb-open="${esc(dataset.id)}">${dataset.delegate ? "APRI CATASTO" : "APRI DATASET"}</button>
      </article>`).join("");
  }

  function currentDataset() { return DATASETS.find((item) => item.id === state.datasetId) || DATASETS[0]; }
  function setStatus(message, type = "") { const node = $("verde-bologna-status"); if (!node) return; node.textContent = message; node.className = `verde-bologna-status ${type}`.trim(); }
  function friendlyLabel(key) { return String(key || "").replace(/^geo_/, "").replace(/_/g, " ").replace(/\b\w/g, (char) => char.toUpperCase()); }

  function isGeometryValue(key, value) {
    const normalized = String(key || "").toLowerCase();
    return normalized.includes("geo_shape") || normalized === "geometry" || normalized === "geom" || normalized.includes("geo_point") || normalized.includes("geopoint") || (value && typeof value === "object" && (value.type || value.geometry || ("lat" in value && ("lon" in value || "lng" in value))));
  }

  function displayValue(value) {
    if (value === null || value === undefined || value === "") return "";
    if (typeof value === "boolean") return value ? "Sì" : "No";
    if (Array.isArray(value)) return value.map(displayValue).filter(Boolean).join(", ");
    if (typeof value === "object") { try { return JSON.stringify(value); } catch (_) { return String(value); } }
    return String(value);
  }

  function primitiveEntries(record) {
    return Object.entries(record || {}).filter(([key, value]) => !isGeometryValue(key, value) && value !== null && value !== undefined && value !== "" && typeof value !== "function").map(([key, value]) => ({ key, label: friendlyLabel(key), value: displayValue(value) })).filter((item) => item.value).slice(0, 10);
  }

  function findField(record, candidates) {
    const entries = Object.entries(record || {});
    for (const candidate of candidates) { const exact = entries.find(([key]) => key.toLowerCase() === candidate); if (exact && displayValue(exact[1])) return displayValue(exact[1]); }
    for (const candidate of candidates) { if (String(candidate).length < 3) continue; const fuzzy = entries.find(([key]) => key.toLowerCase().includes(candidate)); if (fuzzy && displayValue(fuzzy[1])) return displayValue(fuzzy[1]); }
    return "";
  }

  function recordTitle(record, index) { return findField(record, ["denominazione", "nome", "name", "descrizione", "desc", "localizzazione", "specie", "classe", "toponimo", "codice", "id"]) || `${currentDataset().title} · ${index + 1}`; }
  function recordSubtitle(record) { return [findField(record, ["via", "indirizzo", "localita", "quartiere"]), findField(record, ["tipo", "tipologia", "classe", "specie"])].filter(Boolean).filter((value, index, values) => values.indexOf(value) === index).join(" · "); }

  function parseGeoValue(value) {
    if (!value) return null;
    if (typeof value === "string") { try { return parseGeoValue(JSON.parse(value)); } catch (_) { return null; } }
    if (value.type === "Feature" && value.geometry) return value.geometry;
    if (value.geometry) return parseGeoValue(value.geometry);
    if (value.type && Array.isArray(value.coordinates)) return value;
    if (Number.isFinite(Number(value.lat)) && Number.isFinite(Number(value.lon ?? value.lng))) return { type: "Point", coordinates: [Number(value.lon ?? value.lng), Number(value.lat)] };
    return null;
  }

  function plausiblePoint(lon, lat) { return Number.isFinite(lon) && Number.isFinite(lat) && Math.abs(lat) <= 90 && Math.abs(lon) <= 180; }
  function geometryOf(record) {
    for (const [key, value] of Object.entries(record || {})) { if (!isGeometryValue(key, value)) continue; const geometry = parseGeoValue(value); if (geometry) return geometry; }
    const lat = Number(findField(record, ["lat", "latitude", "y"]));
    const lon = Number(findField(record, ["lon", "lng", "longitude", "x"]));
    return plausiblePoint(lon, lat) ? { type: "Point", coordinates: [lon, lat] } : null;
  }

  function flattenCoordinates(coords, output = []) {
    if (!Array.isArray(coords)) return output;
    if (coords.length >= 2 && Number.isFinite(Number(coords[0])) && Number.isFinite(Number(coords[1]))) { output.push([Number(coords[0]), Number(coords[1])]); return output; }
    coords.forEach((item) => flattenCoordinates(item, output)); return output;
  }

  function centerOfGeometry(geometry) {
    if (!geometry) return null;
    const points = flattenCoordinates(geometry.coordinates).filter(([lon, lat]) => plausiblePoint(lon, lat));
    if (!points.length) return null;
    const sum = points.reduce((acc, [lon, lat]) => ({ lon: acc.lon + lon, lat: acc.lat + lat }), { lon: 0, lat: 0 });
    return { lon: sum.lon / points.length, lat: sum.lat / points.length };
  }

  function initializeMap() {
    if (state.map || !window.L || !$("verde-bologna-map")) return;
    state.map = L.map($("verde-bologna-map"), { zoomControl: true, zoomAnimation: false, fadeAnimation: false, markerZoomAnimation: false }).setView([44.4949, 11.3426], 12);
    state.baseLayer = L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "&copy; OpenStreetMap contributors" }).addTo(state.map);
    state.featureLayer = L.layerGroup().addTo(state.map);
  }

  function resizeMap() { requestAnimationFrame(() => state.map?.invalidateSize({ pan: false, animate: false })); setTimeout(() => state.map?.invalidateSize({ pan: false, animate: false }), 180); }

  function addGeometryToMap(record, index, combinedBounds) {
    const geometry = geometryOf(record); if (!geometry || !state.featureLayer) return null;
    const title = recordTitle(record, index); let layer = null;
    try {
      layer = L.geoJSON({ type: "Feature", geometry, properties: {} }, { style: { color: "#187443", weight: 2, fillColor: "#3ba868", fillOpacity: 0.2 }, pointToLayer: (_feature, latlng) => L.circleMarker(latlng, { radius: 7, color: "#fff", weight: 2, fillColor: "#18854b", fillOpacity: 0.95 }) }).addTo(state.featureLayer);
      layer.bindPopup(`<strong>${esc(title)}</strong><br><span>Comune di Bologna Open Data</span>`);
      const bounds = layer.getBounds?.(); if (bounds?.isValid?.()) combinedBounds.extend(bounds);
      const center = centerOfGeometry(geometry); if (center) combinedBounds.extend([center.lat, center.lon]);
      state.featureByIndex.set(index, { layer, center }); return center;
    } catch (_) { return null; }
  }

  function renderMap() {
    initializeMap(); state.featureLayer?.clearLayers(); state.featureByIndex.clear();
    const combinedBounds = L.latLngBounds([]); let geocoded = 0;
    state.records.forEach((record, index) => { if (addGeometryToMap(record, index, combinedBounds)) geocoded += 1; });
    if (combinedBounds.isValid()) state.map.fitBounds(combinedBounds.pad(0.08), { animate: false, maxZoom: 16 }); else state.map.setView([44.4949, 11.3426], 12, { animate: false });
    const mapStatus = $("verde-bologna-map-status"); if (mapStatus) mapStatus.textContent = geocoded ? `${geocoded} elementi dei ${state.records.length} caricati hanno geometria utilizzabile sulla mappa.` : "I record caricati non espongono coordinate utilizzabili in questa pagina; i dati testuali restano consultabili.";
    resizeMap();
  }

  function focusResult(index) {
    const item = state.featureByIndex.get(index); if (!item || !state.map) return;
    const bounds = item.layer?.getBounds?.(); if (bounds?.isValid?.()) state.map.fitBounds(bounds.pad(0.2), { animate: false, maxZoom: 18 }); else if (item.center) state.map.setView([item.center.lat, item.center.lon], 18, { animate: false });
    try { const layers = item.layer?.getLayers?.() || []; (layers[0] || item.layer)?.openPopup?.(); } catch (_) {}
    $("verde-bologna-map-card")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function renderResults() {
    const node = $("verde-bologna-results"); if (!node) return;
    if (!state.records.length) { node.innerHTML = `<p class="verde-bologna-empty">Nessun record trovato con i filtri attuali.</p>`; $("verde-bologna-load-more")?.classList.add("hidden"); renderMap(); return; }
    node.innerHTML = state.records.map((record, index) => {
      const title = recordTitle(record, index), subtitle = recordSubtitle(record), entries = primitiveEntries(record), center = centerOfGeometry(geometryOf(record));
      const navHref = center ? `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${center.lat},${center.lon}`)}` : "";
      return `<article class="verde-bologna-result"><div><h3>${esc(title)}</h3>${subtitle ? `<p>${esc(subtitle)}</p>` : ""}</div><div class="verde-bologna-result-actions">${center ? `<button class="btn" type="button" data-vb-map-index="${index}">MOSTRA</button><a class="btn btn-primary" href="${esc(navHref)}" target="_blank" rel="noopener">NAVIGA</a>` : ""}</div><div class="verde-bologna-details">${entries.map((entry) => `<div><span>${esc(entry.label)}</span><strong>${esc(entry.value)}</strong></div>`).join("")}</div></article>`;
    }).join("");
    node.querySelectorAll("[data-vb-map-index]").forEach((button) => button.addEventListener("click", () => focusResult(Number(button.dataset.vbMapIndex))));
    const more = $("verde-bologna-load-more"); more?.classList.toggle("hidden", state.records.length >= state.total || state.records.length === 0); if (more && !more.classList.contains("hidden")) more.textContent = `CARICA ALTRI 100 · ${state.records.length} / ${state.total}`;
    renderMap();
  }

  function localMatches(record, query) {
    const needle = String(query || "").trim().toLocaleLowerCase("it-IT"); if (!needle) return true;
    return Object.entries(record || {}).some(([key, value]) => !isGeometryValue(key, value) && displayValue(value).toLocaleLowerCase("it-IT").includes(needle));
  }

  async function requestRecords(datasetId, offset, query) {
    const attempt = async (includeSearch) => {
      const cached = readCache(datasetId, offset, query, includeSearch); if (cached) return { payload: cached, includeSearch };
      const response = await fetch(apiUrl(datasetId, offset, query, includeSearch), { headers: { Accept: "application/json" } });
      if (!response.ok) throw new Error(`API Comune di Bologna non disponibile (${response.status}).`);
      const payload = await response.json(); writeCache(datasetId, offset, query, includeSearch, payload); return { payload, includeSearch };
    };
    if (!query) return attempt(false);
    try { return await attempt(true); }
    catch (_) {
      const fallback = await attempt(false); const records = Array.isArray(fallback.payload?.results) ? fallback.payload.results.filter((record) => localMatches(record, query)) : [];
      return { payload: { ...fallback.payload, total_count: records.length, results: records }, includeSearch: false, localFallback: true };
    }
  }

  async function loadRecords({ append = false } = {}) {
    const dataset = currentDataset(); if (dataset.delegate) return;
    if (navigator.onLine === false) { setStatus("Questa sezione richiede Internet per interrogare gli Open Data del Comune di Bologna.", "error"); return; }
    const serial = ++state.requestSerial, offset = append ? state.records.length : 0;
    if (!append) { state.offset = 0; state.records = []; state.total = 0; }
    setStatus(`${append ? "Carico altri record" : "Interrogo il dataset ufficiale"} “${dataset.title}”…`);
    const more = $("verde-bologna-load-more"); if (more) more.disabled = true;
    try {
      const response = await requestRecords(dataset.id, offset, state.query); if (serial !== state.requestSerial) return;
      const results = Array.isArray(response.payload?.results) ? response.payload.results : [];
      state.records = append ? [...state.records, ...results] : results; state.total = Number(response.payload?.total_count ?? state.records.length) || state.records.length; state.offset = state.records.length;
      renderResults(); const fallbackNote = response.localFallback ? " Ricerca applicata ai 100 record della pagina perché il filtro testuale remoto non era disponibile." : "";
      setStatus(`${state.records.length} record caricati${state.total ? ` su ${state.total}` : ""}.${fallbackNote}`, response.localFallback ? "warning" : "");
    } catch (error) { if (serial !== state.requestSerial) return; setStatus(error?.message || "Impossibile leggere il dataset del Comune di Bologna.", "error"); }
    finally { if (more) more.disabled = false; }
  }

  function openDataset(datasetId) {
    const dataset = DATASETS.find((item) => item.id === datasetId); if (!dataset) return;
    if (dataset.delegate) { closePage(); const target = $(dataset.delegate); if (target) target.click(); else window.alert("Catasto alberi non disponibile in questo momento."); return; }
    state.datasetId = dataset.id; state.query = ""; state.records = []; state.total = 0; renderDatasetCards(); $("verde-bologna-browser")?.classList.remove("hidden");
    const title = $("verde-bologna-active-title"), description = $("verde-bologna-active-description"), source = $("verde-bologna-source-link"), query = $("verde-bologna-query");
    if (title) title.textContent = `${dataset.icon} ${dataset.title}`; if (description) description.textContent = dataset.short; if (source) source.href = sourcePageUrl(dataset.id); if (query) query.value = "";
    $("verde-bologna-browser")?.scrollIntoView({ behavior: "smooth", block: "start" }); loadRecords();
  }

  function openPage() {
    $("menu-close-btn")?.click(); $("home-page")?.classList.add("hidden"); const page = buildPage(); page.classList.remove("hidden"); page.setAttribute("aria-hidden", "false"); renderDatasetCards(); initializeMap(); resizeMap(); if (!state.records.length) openDataset(state.datasetId || "un_gest");
  }

  function closePage() { setFullscreen(false); const page = $(PAGE_ID); page?.classList.add("hidden"); page?.setAttribute("aria-hidden", "true"); $("home-page")?.classList.remove("hidden"); }

  function setFullscreen(active) {
    state.fullscreen = Boolean(active); $("verde-bologna-map-card")?.classList.toggle("is-fullscreen", state.fullscreen); document.body.classList.toggle("verde-bologna-fullscreen-open", state.fullscreen);
    const button = $("verde-bologna-fullscreen-btn"); if (button) { button.setAttribute("aria-pressed", String(state.fullscreen)); button.textContent = state.fullscreen ? "✕ CHIUDI MAPPA" : "⛶ SCHERMO INTERO"; } resizeMap();
  }

  function showUserLocation() {
    if (!navigator.geolocation) { setStatus("Geolocalizzazione non supportata da questo dispositivo.", "error"); return; }
    const button = $("verde-bologna-location-btn"); if (button) button.disabled = true;
    navigator.geolocation.getCurrentPosition((position) => {
      initializeMap(); const lat = Number(position.coords.latitude), lon = Number(position.coords.longitude), accuracy = Number(position.coords.accuracy); if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
      const point = L.latLng(lat, lon); if (!state.userMarker) state.userMarker = L.circleMarker(point, { radius: 9, color: "#fff", weight: 3, fillColor: "#1268e8", fillOpacity: 1 }).addTo(state.map).bindPopup("<strong>La mia posizione</strong>"); else state.userMarker.setLatLng(point).addTo(state.map);
      if (Number.isFinite(accuracy) && accuracy > 0) { if (!state.userAccuracy) state.userAccuracy = L.circle(point, { radius: accuracy, color: "#1268e8", weight: 1, fillOpacity: 0.08 }).addTo(state.map); else state.userAccuracy.setLatLng(point).setRadius(accuracy).addTo(state.map); }
      state.map.setView(point, Math.max(state.map.getZoom(), 16), { animate: false }); state.userMarker.openPopup(); if (button) button.disabled = false;
    }, (error) => { if (button) button.disabled = false; setStatus(error?.code === 1 ? "Permesso posizione negato." : "Impossibile trovare la posizione GPS.", "error"); }, { enableHighAccuracy: true, timeout: 12000, maximumAge: 30000 });
  }

  function addMenuButton() {
    if ($(MENU_BUTTON_ID)) return;
    const anchor = $("open-green-areas-btn") || $("open-tree-search-btn") || document.querySelector("#side-menu .menu-section"); if (!anchor) return;
    const button = document.createElement("button"); button.id = MENU_BUTTON_ID; button.className = "btn menu-title-btn"; button.type = "button"; button.innerHTML = '<span class="menu-item-icon" aria-hidden="true">🌳</span>Verde Bologna';
    if (anchor.matches?.("button")) anchor.insertAdjacentElement("beforebegin", button); else anchor.appendChild(button); button.addEventListener("click", openPage);
  }

  function installEvents() {
    $("verde-bologna-back-btn")?.addEventListener("click", closePage);
    $("verde-bologna-datasets")?.addEventListener("click", (event) => { const button = event.target.closest("[data-vb-open]"); if (button) openDataset(button.dataset.vbOpen); });
    $("verde-bologna-search-form")?.addEventListener("submit", (event) => { event.preventDefault(); state.query = $("verde-bologna-query")?.value.trim() || ""; loadRecords(); });
    $("verde-bologna-clear-btn")?.addEventListener("click", () => { state.query = ""; if ($("verde-bologna-query")) $("verde-bologna-query").value = ""; loadRecords(); });
    $("verde-bologna-load-more")?.addEventListener("click", () => loadRecords({ append: true })); $("verde-bologna-location-btn")?.addEventListener("click", showUserLocation); $("verde-bologna-fullscreen-btn")?.addEventListener("click", () => setFullscreen(!state.fullscreen));
    document.addEventListener("keydown", (event) => { if (event.key !== "Escape") return; if (state.fullscreen) setFullscreen(false); else if (!$(PAGE_ID)?.classList.contains("hidden")) closePage(); });
  }

  function install() {
    injectStyle(); buildPage(); renderDatasetCards(); addMenuButton(); installEvents();
    window.HeraVerdeBologna = Object.freeze({ open: openPage, close: closePage, openDataset, datasets: DATASETS.map(({ id, title }) => ({ id, title })) });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true }); else install();
})();

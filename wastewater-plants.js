(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const page = $("wastewater-plants-page");
  const form = $("wastewater-plants-form");
  const queryInput = $("wastewater-plants-query");
  const kindInput = $("wastewater-plants-kind");
  const provinceInput = $("wastewater-plants-province");
  const statusNode = $("wastewater-plants-status");
  const resultsNode = $("wastewater-plants-results");
  const mapCard = $("wastewater-plants-map-card");
  const mapNode = $("wastewater-plants-map");
  const mapStatus = $("wastewater-plants-map-status");
  const mapStyle = $("wastewater-plants-map-style");
  const fullscreenButton = $("wastewater-plants-fullscreen-btn");
  const sheet = $("wastewater-plant-sheet");
  const sheetTitle = $("wastewater-plant-sheet-title");
  const sheetSource = $("wastewater-plant-sheet-source");
  const sheetBody = $("wastewater-plant-sheet-body");
  const navigateButton = $("wastewater-plant-navigate");
  const API_URL = "https://servizi-gis.arpae.it/server/rest/services/Geoportal/ACQUEPressioni/MapServer/1/query";
  const SOURCE_URL = "https://dati.arpae.it/dataset/arpa_acq_reflue_urbane_depurate_depurat_tutti_22_e23";
  const OSM_SOURCE_URL = "https://www.openstreetmap.org/copyright";
  const LIFT_API_URL = window.Capacitor?.isNativePlatform?.()
    ? "https://creative-syrniki-dddbae.netlify.app/api/wastewater-lift-stations"
    : "/api/wastewater-lift-stations";
  const CACHE_KEY = "varga-arpae-wastewater-plants:v1";
  const CACHE_TTL_MS = 30 * 60 * 1000;
  const PAGE_SIZE = 1000;
  const RESULT_LIST_LIMIT = 100;
  const MAX_RECORDS_SAFETY = 5000;

  const TILE_LAYERS = Object.freeze({
    classic: { url: "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "&copy; OpenStreetMap contributors" } },
    satellite: { url: "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "Tiles &copy; Esri" } },
    labels: { url: "https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "Labels &copy; Esri", pane: "overlayPane" } }
  });

  let map = null;
  let baseLayer = null;
  let hybridLabels = null;
  let markersLayer = null;
  let operatorMarker = null;
  let fullscreen = false;
  let allPlants = [];
  const liftStationsByArea = new Map();
  let currentItems = [];
  let selectedItem = null;
  let loadingPromise = null;

  const esc = (value) => String(value ?? "—").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char]));
  const clean = (value) => String(value ?? "").replace(/\s+/g, " ").trim();
  const upper = (value) => clean(value).toLocaleUpperCase("it-IT");
  const formatNumber = (value) => Number.isFinite(Number(value)) ? new Intl.NumberFormat("it-IT").format(Number(value)) : "Non disponibile";

  function setStatus(message, type = "") {
    if (!statusNode) return;
    statusNode.textContent = message;
    statusNode.className = `wastewater-plants-status ${type}`.trim();
  }

  function applyMapStyle(style) {
    if (!map) return;
    baseLayer?.remove();
    hybridLabels?.remove();
    const selected = style === "satellite" || style === "hybrid" ? TILE_LAYERS.satellite : TILE_LAYERS.classic;
    baseLayer = L.tileLayer(selected.url, selected.options).addTo(map);
    baseLayer.bringToBack();
    if (style === "hybrid") hybridLabels = L.tileLayer(TILE_LAYERS.labels.url, TILE_LAYERS.labels.options).addTo(map);
  }

  function initializeMap() {
    if (map || !window.L || !mapNode) return;
    map = L.map(mapNode, { zoomControl: true, zoomAnimation: false, fadeAnimation: false, markerZoomAnimation: false }).setView([44.55, 11.1], 8);
    markersLayer = L.layerGroup().addTo(map);
    applyMapStyle(mapStyle?.value || "classic");
  }

  function resizeMap() {
    requestAnimationFrame(() => map?.invalidateSize({ pan: false, animate: false }));
    setTimeout(() => map?.invalidateSize({ pan: false, animate: false }), 180);
  }

  function setFullscreen(active) {
    fullscreen = Boolean(active);
    mapCard?.classList.toggle("wastewater-plants-map-card--fullscreen", fullscreen);
    document.body.classList.toggle("wastewater-plants-fullscreen-open", fullscreen);
    fullscreenButton?.setAttribute("aria-pressed", String(fullscreen));
    if (fullscreenButton) fullscreenButton.textContent = fullscreen ? "✕ CHIUDI MAPPA" : "⛶ SCHERMO INTERO";
    resizeMap();
  }

  function readCache() {
    try {
      const cached = JSON.parse(sessionStorage.getItem(CACHE_KEY) || "null");
      return cached && Date.now() - cached.savedAt < CACHE_TTL_MS && Array.isArray(cached.items) ? cached.items : null;
    } catch (_) { return null; }
  }

  function writeCache(items) {
    try { sessionStorage.setItem(CACHE_KEY, JSON.stringify({ savedAt: Date.now(), items })); } catch (_) {}
  }

  function normalizeFeature(feature) {
    const properties = feature?.properties || {};
    const coordinates = feature?.geometry?.coordinates || [];
    const lon = Number(coordinates[0]);
    const lat = Number(coordinates[1]);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    return {
      id: clean(properties.OBJECTID || feature.id),
      code: clean(properties.COD_DEP),
      name: clean(properties.DEN_DEP) || "Depuratore senza denominazione",
      municipality: clean(properties.COMUNE),
      municipalityCode: clean(properties.ISTAT_COD),
      province: clean(properties.NOME_PROV),
      provinceCode: clean(properties.COD_PROV),
      provinceAbbreviation: clean(properties.SIGLA),
      manager: clean(properties.GESTORE),
      typeCode: clean(properties.TIPO_DEP),
      typeDescription: clean(properties.DESCR_TIPO),
      treatmentLevel: clean(properties.LIV_DEP),
      designPopulation: Number(properties.AE_PROG),
      agglomerationCode: clean(properties.COD_AGG),
      agglomerationName: clean(properties.NOME_AGG),
      agglomerationClass: clean(properties.CLASSE_AGG),
      waterBody: clean(properties.N__CIS_WFD),
      waterBodyCode: clean(properties.COD_CIS),
      waterBodyDetail: clean(properties.N_CIS),
      kind: "depuratore",
      source: "ARPAE",
      lat,
      lon
    };
  }

  const OSM_AREA_NAMES = Object.freeze({
    "": "Emilia-Romagna",
    BOLOGNA: "Bologna",
    FERRARA: "Ferrara",
    "FORLI'-CESENA": "Forlì-Cesena",
    MODENA: "Modena",
    PARMA: "Parma",
    PIACENZA: "Piacenza",
    RAVENNA: "Ravenna",
    "REGGIO NELL'EMILIA": "Reggio Emilia",
    RIMINI: "Rimini"
  });

  function normalizeLiftStation(element, area) {
    const tags = element?.tags || {};
    const lat = Number(element?.lat ?? element?.center?.lat);
    const lon = Number(element?.lon ?? element?.center?.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    const osmId = `${element.type || "node"}/${element.id}`;
    const declaredType = upper(`${tags.pumping_station || ""} ${tags.substance || ""} ${tags.name || ""}`);
    const confirmedSewage = /(SEWAGE|WASTEWATER|FOGN|REFLU|SOLLEV)/.test(declaredType);
    return {
      id: `osm-${osmId}`,
      code: clean(tags.ref || tags.operator_ref || `OSM ${element.id}`),
      name: clean(tags.name || tags.local_ref) || (confirmedSewage ? "Sollevamento fognario" : "Stazione di pompaggio"),
      municipality: clean(tags["addr:city"] || tags["addr:place"]),
      province: area === "Emilia-Romagna" ? "" : upper(area === "Reggio Emilia" ? "Reggio nell'Emilia" : area),
      provinceAbbreviation: "OSM",
      manager: clean(tags.operator || tags.owner),
      typeDescription: confirmedSewage ? "Impianto di sollevamento fognario" : "Stazione di pompaggio · tipo non specificato su OSM",
      confirmedSewage,
      pumpingStation: clean(tags.pumping_station),
      substance: clean(tags.substance),
      address: clean([tags["addr:street"], tags["addr:housenumber"]].filter(Boolean).join(" ")),
      access: clean(tags.access),
      website: clean(tags.website),
      osmType: element.type || "node",
      osmNumericId: element.id,
      osmUrl: `https://www.openstreetmap.org/${osmId}`,
      kind: "sollevamento",
      source: "OpenStreetMap",
      lat,
      lon
    };
  }

  async function loadLiftStations() {
    const area = OSM_AREA_NAMES[provinceInput?.value || ""] || "Emilia-Romagna";
    if (liftStationsByArea.has(area)) return liftStationsByArea.get(area);
    setStatus(`Caricamento sollevamenti fognari OpenStreetMap · ${area}…`);
    const response = await fetch(`${LIFT_API_URL}?area=${encodeURIComponent(area)}`, { headers: { Accept: "application/json" } });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok || !payload.ok || !Array.isArray(payload.elements)) throw new Error(payload.error || `OpenStreetMap non disponibile (${response.status}).`);
    const seen = new Set();
    const items = payload.elements.map((element) => normalizeLiftStation(element, area)).filter(Boolean).filter((item) => {
      if (seen.has(item.id)) return false;
      seen.add(item.id);
      return true;
    });
    liftStationsByArea.set(area, items);
    return items;
  }

  async function fetchPage(offset) {
    const params = new URLSearchParams({
      where: "1=1",
      outFields: "*",
      returnGeometry: "true",
      outSR: "4326",
      orderByFields: "OBJECTID",
      resultOffset: String(offset),
      resultRecordCount: String(PAGE_SIZE),
      f: "geojson"
    });
    const response = await fetch(`${API_URL}?${params}`, { headers: { Accept: "application/geo+json, application/json" } });
    if (!response.ok) throw new Error(`API ARPAE non disponibile (${response.status}).`);
    const payload = await response.json();
    if (!Array.isArray(payload.features)) throw new Error("Risposta ARPAE non valida.");
    return payload;
  }

  async function loadAllPlants() {
    if (allPlants.length) return allPlants;
    const cached = readCache();
    if (cached?.length) {
      allPlants = cached;
      return allPlants;
    }
    if (loadingPromise) return loadingPromise;
    loadingPromise = (async () => {
      const features = [];
      for (let offset = 0; offset < MAX_RECORDS_SAFETY; offset += PAGE_SIZE) {
        setStatus(`Caricamento ARPAE… ${offset ? `${offset} impianti acquisiti` : "prima pagina"}.`);
        const payload = await fetchPage(offset);
        features.push(...payload.features);
        if (!payload.exceededTransferLimit && payload.features.length < PAGE_SIZE) break;
      }
      const seen = new Set();
      allPlants = features.map(normalizeFeature).filter(Boolean).filter((item) => {
        const key = item.code || item.id;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
      if (!allPlants.length) throw new Error("Il censimento ARPAE non contiene impianti utilizzabili.");
      writeCache(allPlants);
      return allPlants;
    })().finally(() => { loadingPromise = null; });
    return loadingPromise;
  }

  function filterPlants() {
    const needle = upper(queryInput?.value);
    const province = upper(provinceInput?.value);
    const kind = kindInput?.value || "all";
    const area = OSM_AREA_NAMES[provinceInput?.value || ""] || "Emilia-Romagna";
    const liftStations = liftStationsByArea.get(area) || [];
    return [...allPlants, ...liftStations].filter((item) => {
      if (kind !== "all" && item.kind !== kind) return false;
      if (item.kind === "depuratore" && province && upper(item.province) !== province) return false;
      if (!needle) return true;
      const haystack = upper([
        item.code,
        item.name,
        item.municipality,
        item.manager,
        item.agglomerationCode,
        item.agglomerationName,
        item.typeDescription,
        item.waterBody,
        item.address,
        item.access,
        item.source
      ].join(" "));
      return haystack.includes(needle);
    });
  }

  function markerColor(item) {
    if (item.kind === "sollevamento") return "#7a3db8";
    const capacity = Number(item.designPopulation) || 0;
    if (capacity >= 10000) return "#b42318";
    if (capacity >= 2000) return "#f79009";
    if (capacity >= 200) return "#198754";
    return "#1570ef";
  }

  function showItems(items) {
    initializeMap();
    markersLayer?.clearLayers();
    currentItems = items;
    const bounds = L.latLngBounds([]);
    items.forEach((item, index) => {
      L.circleMarker([item.lat, item.lon], {
        radius: item.kind === "sollevamento" || Number(item.designPopulation) >= 10000 ? 8 : 6,
        color: "#fff",
        weight: 1.5,
        fillColor: markerColor(item),
        fillOpacity: 0.9
      }).addTo(markersLayer).bindPopup(`<strong>${esc(item.name)}</strong><br>${esc(item.kind === "sollevamento" ? `${item.typeDescription} · OpenStreetMap` : `${item.code || "Codice non disponibile"} · ${item.municipality || "Comune non disponibile"}`)}${item.kind === "depuratore" ? `<br>${esc(formatNumber(item.designPopulation))} A.E.` : ""}<br><button class="btn btn-primary" type="button" data-wastewater-plant-index="${index}">APRI SCHEDA</button>`);
      bounds.extend([item.lat, item.lon]);
    });
    if (bounds.isValid()) map.fitBounds(bounds.pad(0.08), { animate: false, maxZoom: 16 });
    else map.setView([44.55, 11.1], 8, { animate: false });
    const lifts = items.filter((item) => item.kind === "sollevamento").length;
    const plants = items.length - lifts;
    mapStatus.textContent = `${items.length} impianti visualizzati · ${plants} depuratori · ${lifts} stazioni OSM.`;
    resizeMap();
  }

  function renderResults(items) {
    const shown = items.slice(0, RESULT_LIST_LIMIT);
    resultsNode.innerHTML = shown.map((item, index) => `<article class="wastewater-plant-result">
      <div>${item.kind === "sollevamento" ? `<span class="wastewater-plant-source-badge">${item.confirmedSewage ? "SOLLEVAMENTO FOGNARIO" : "POMPA · TIPO NON SPECIFICATO"} · OSM</span>` : ""}<small>${esc(item.code || "Senza codice")} · ${esc(item.provinceAbbreviation || item.province)}</small><h2>${esc(item.name)}</h2><p>${item.kind === "sollevamento" ? `${esc(item.municipality || item.address || "Località non indicata")} · ${esc(item.manager || "Gestore non indicato")}` : `${esc(item.municipality || "Comune non disponibile")} · ${esc(item.manager || "Gestore non disponibile")} · ${esc(formatNumber(item.designPopulation))} A.E.`}</p></div>
      <button class="btn" type="button" data-wastewater-result-index="${index}">MOSTRA E APRI SCHEDA</button>
    </article>`).join("");
    if (items.length > RESULT_LIST_LIMIT) {
      resultsNode.insertAdjacentHTML("beforeend", `<p class="wastewater-plant-result wastewater-plants-list-note">Nell’elenco sono mostrati i primi ${RESULT_LIST_LIMIT} risultati; tutti i ${items.length} impianti sono presenti sulla mappa. Usa nome, codice, Comune o gestore per restringere la ricerca.</p>`);
    }
    resultsNode.classList.toggle("hidden", !items.length);
  }

  function detailRow(label, value) {
    const text = clean(value);
    return text ? `<dt>${esc(label)}</dt><dd>${esc(text)}</dd>` : "";
  }

  function openSheet(item) {
    if (!item) return;
    selectedItem = item;
    sheetTitle.textContent = item.name;
    if (sheetSource) sheetSource.textContent = item.kind === "sollevamento" ? "Fonte: OpenStreetMap · dato collaborativo" : "Fonte: ARPAE Emilia-Romagna · edizione 2023";
    if (item.kind === "sollevamento") {
      sheetBody.innerHTML = `<section class="wastewater-plant-detail-section"><h3>Identificazione</h3><dl>
        ${detailRow("Tipo", item.typeDescription)}
        ${detailRow("Codice / riferimento", item.code)}
        ${detailRow("Comune o località", item.municipality)}
        ${detailRow("Indirizzo", item.address)}
        ${detailRow("Gestore", item.manager)}
        ${detailRow("Accesso", item.access)}
        ${detailRow("Fluido indicato", item.substance || item.pumpingStation)}
        ${detailRow("Coordinate", `${item.lat.toFixed(6)}, ${item.lon.toFixed(6)}`)}
      </dl></section>
      <p class="wastewater-plants-list-note">Dato collaborativo OpenStreetMap: quando il tipo non è specificato, non è possibile confermare che la stazione appartenga alla rete fognaria.</p>
      <a class="wastewater-plant-source-link" href="${esc(item.osmUrl || OSM_SOURCE_URL)}" target="_blank" rel="noopener">Apri l’elemento su OpenStreetMap</a>`;
    } else {
      sheetBody.innerHTML = `<section class="wastewater-plant-detail-section"><h3>Identificazione</h3><dl>
      ${detailRow("Codice depuratore", item.code)}
      ${detailRow("Comune", item.municipality)}
      ${detailRow("Provincia", `${item.province}${item.provinceAbbreviation ? ` (${item.provinceAbbreviation})` : ""}`)}
      ${detailRow("Codice ISTAT Comune", item.municipalityCode)}
      ${detailRow("Gestore", item.manager)}
    </dl></section>
    <section class="wastewater-plant-detail-section"><h3>Trattamento</h3><dl>
      ${detailRow("Tipo", item.typeDescription || item.typeCode)}
      ${detailRow("Codice tipo", item.typeCode)}
      ${detailRow("Livello depurazione", item.treatmentLevel)}
      ${detailRow("Capacità di progetto", `${formatNumber(item.designPopulation)} A.E.`)}
      ${detailRow("Classe agglomerato", item.agglomerationClass)}
    </dl></section>
    <section class="wastewater-plant-detail-section"><h3>Agglomerato e corpo idrico</h3><dl>
      ${detailRow("Agglomerato", item.agglomerationName)}
      ${detailRow("Codice agglomerato", item.agglomerationCode)}
      ${detailRow("Corpo idrico", item.waterBody)}
      ${detailRow("Dettaglio corpo idrico", item.waterBodyDetail)}
      ${detailRow("Codice corpo idrico", item.waterBodyCode)}
      ${detailRow("Coordinate", `${item.lat.toFixed(6)}, ${item.lon.toFixed(6)}`)}
    </dl></section>
      <a class="wastewater-plant-source-link" href="${SOURCE_URL}" target="_blank" rel="noopener">Apri il dataset ufficiale ARPAE</a>`;
    }
    navigateButton.href = `https://www.google.com/maps/dir/?api=1&destination=${item.lat},${item.lon}`;
    sheet.classList.remove("hidden");
    sheet.setAttribute("aria-hidden", "false");
  }

  function closeSheet() {
    sheet?.classList.add("hidden");
    sheet?.setAttribute("aria-hidden", "true");
    selectedItem = null;
  }

  async function openStreetView() {
    if (!selectedItem) return;
    const api = window.HeraStreetViewCards;
    if (typeof api?.openForCoordinates !== "function") return window.alert("Vista 360° non disponibile in questo momento. Riprova tra qualche secondo.");
    await api.openForCoordinates({ lat: selectedItem.lat, lng: selectedItem.lon }, $("wastewater-plant-street-view"), {
      targetLabel: selectedItem.name,
      modalTitle: `🌐 Vista 360° ${selectedItem.name}`
    });
  }

  async function shareSelectedItem() {
    if (!selectedItem) return;
    const navigationUrl = `https://www.google.com/maps/dir/?api=1&destination=${selectedItem.lat},${selectedItem.lon}`;
    const isLift = selectedItem.kind === "sollevamento";
    const message = [
      isLift ? (selectedItem.confirmedSewage ? "🟣 *SOLLEVAMENTO FOGNARIO*" : "🟣 *STAZIONE DI POMPAGGIO OSM*") : "🏭 *SCHEDA DEPURATORE*",
      "",
      `• *Impianto:* ${selectedItem.name}`,
      `• *Codice:* ${selectedItem.code || "Non disponibile"}`,
      `• *Comune:* ${selectedItem.municipality || "Non disponibile"}`,
      `• *Gestore:* ${selectedItem.manager || "Non disponibile"}`,
      `• *Tipo:* ${selectedItem.typeDescription || selectedItem.typeCode || "Non disponibile"}`,
      ...(isLift ? [`• *Fonte:* OpenStreetMap (dato collaborativo)`] : [`• *Capacità:* ${formatNumber(selectedItem.designPopulation)} A.E.`]),
      "",
      isLift ? "📍 *NAVIGA VERSO IL SOLLEVAMENTO*" : "📍 *NAVIGA VERSO IL DEPURATORE*",
      navigationUrl
    ].join("\n");
    const appUrl = `whatsapp://send?text=${encodeURIComponent(message)}`;
    const nativeAndroid = Boolean(window.Capacitor?.isNativePlatform?.() && window.Capacitor?.getPlatform?.() === "android");
    if (nativeAndroid) {
      const plugin = window.Capacitor?.Plugins?.HeraWhatsApp || window.Capacitor?.registerPlugin?.("HeraWhatsApp") || null;
      if (!plugin?.open) return window.alert("WhatsApp non è disponibile su questo dispositivo.");
      try { await plugin.open({ url: appUrl }); } catch (error) { window.alert(error?.message || "WhatsApp non è installato o non può essere aperto."); }
      return;
    }
    window.location.assign(appUrl);
    window.setTimeout(() => { if (document.visibilityState === "visible") window.alert("WhatsApp non è installato o non può essere aperto."); }, 1800);
  }

  function runSearch(loadWarning = "") {
    const items = filterPlants();
    renderResults(items);
    showItems(items);
    if (!items.length) return setStatus(loadWarning || "Nessun impianto corrisponde ai filtri. Prova un nome più breve o cambia provincia.", "error");
    const lifts = items.filter((item) => item.kind === "sollevamento").length;
    const plants = items.length - lifts;
    setStatus(`${items.length} impianti trovati · ${plants} depuratori ARPAE · ${lifts} stazioni OSM.${loadWarning ? ` ${loadWarning}` : ""}`, loadWarning ? "" : "success");
  }

  async function refreshSearch() {
    let warning = "";
    try {
      await loadAllPlants();
      if ((kindInput?.value || "all") !== "depuratore") await loadLiftStations();
    } catch (error) {
      warning = `${error.message || "Fonte supplementare non disponibile"} I depuratori ARPAE restano consultabili.`;
    }
    runSearch(warning);
  }

  async function openPage() {
    $("menu-close-btn")?.click();
    $("home-page")?.classList.add("hidden");
    page?.classList.remove("hidden");
    page?.setAttribute("aria-hidden", "false");
    initializeMap();
    resizeMap();
    try {
      await refreshSearch();
      queryInput?.focus();
    } catch (error) {
      setStatus(`${error.message || "Caricamento non riuscito."} Controlla la connessione e riprova.`, "error");
    }
  }

  function closePage() {
    setFullscreen(false);
    closeSheet();
    page?.classList.add("hidden");
    page?.setAttribute("aria-hidden", "true");
    $("home-page")?.classList.remove("hidden");
  }

  function locateOperator() {
    if (!navigator.geolocation) return setStatus("La posizione non è disponibile su questo dispositivo.", "error");
    setStatus("Rilevamento della posizione dell’operatore…");
    navigator.geolocation.getCurrentPosition(({ coords }) => {
      initializeMap();
      const center = [Number(coords.latitude), Number(coords.longitude)];
      operatorMarker?.remove();
      operatorMarker = L.circleMarker(center, { radius: 8, color: "#fff", weight: 3, fillColor: "#1769e0", fillOpacity: 1 }).addTo(map).bindPopup("La mia posizione").openPopup();
      map.setView(center, Math.max(map.getZoom(), 13), { animate: false });
      setStatus("Posizione mostrata sulla mappa.", "success");
    }, (error) => setStatus(error.code === 1 ? "Permesso posizione negato. Abilitalo nelle impostazioni." : "Non riesco a rilevare la posizione.", "error"), { enableHighAccuracy: true, timeout: 12000, maximumAge: 60000 });
  }

  form?.addEventListener("submit", (event) => { event.preventDefault(); refreshSearch(); });
  provinceInput?.addEventListener("change", refreshSearch);
  kindInput?.addEventListener("change", refreshSearch);
  $("wastewater-plants-clear-btn")?.addEventListener("click", () => {
    queryInput.value = "";
    provinceInput.value = "";
    kindInput.value = "all";
    refreshSearch();
  });
  $("open-wastewater-plants-btn")?.addEventListener("click", openPage);
  $("wastewater-plants-back-btn")?.addEventListener("click", closePage);
  $("wastewater-plants-location-btn")?.addEventListener("click", locateOperator);
  fullscreenButton?.addEventListener("click", () => setFullscreen(!fullscreen));
  mapStyle?.addEventListener("change", () => applyMapStyle(mapStyle.value));
  resultsNode?.addEventListener("click", (event) => {
    const button = event.target.closest?.("[data-wastewater-result-index]");
    if (!button) return;
    const item = currentItems[Number(button.dataset.wastewaterResultIndex)];
    map?.setView([item.lat, item.lon], 17, { animate: false });
    openSheet(item);
  });
  mapNode?.addEventListener("click", (event) => {
    const button = event.target.closest?.("[data-wastewater-plant-index]");
    if (button) openSheet(currentItems[Number(button.dataset.wastewaterPlantIndex)]);
  });
  $("wastewater-plant-sheet-close")?.addEventListener("click", closeSheet);
  $("wastewater-plant-street-view")?.addEventListener("click", openStreetView);
  $("wastewater-plant-whazzup")?.addEventListener("click", shareSelectedItem);
  sheet?.addEventListener("click", (event) => { if (event.target === sheet) closeSheet(); });
  document.addEventListener("keydown", (event) => { if (event.key === "Escape") fullscreen ? setFullscreen(false) : closeSheet(); });
})();

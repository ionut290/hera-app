(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const page = $("wastewater-plants-page");
  const form = $("wastewater-plants-form");
  const queryInput = $("wastewater-plants-query");
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
  const sheetBody = $("wastewater-plant-sheet-body");
  const navigateButton = $("wastewater-plant-navigate");
  const API_URL = "https://servizi-gis.arpae.it/server/rest/services/Geoportal/ACQUEPressioni/MapServer/1/query";
  const SOURCE_URL = "https://dati.arpae.it/dataset/arpa_acq_reflue_urbane_depurate_depurat_tutti_22_e23";
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
      lat,
      lon
    };
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
    return allPlants.filter((item) => {
      if (province && upper(item.province) !== province) return false;
      if (!needle) return true;
      const haystack = upper([
        item.code,
        item.name,
        item.municipality,
        item.manager,
        item.agglomerationCode,
        item.agglomerationName,
        item.typeDescription,
        item.waterBody
      ].join(" "));
      return haystack.includes(needle);
    });
  }

  function markerColor(item) {
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
        radius: Number(item.designPopulation) >= 10000 ? 8 : 6,
        color: "#fff",
        weight: 1.5,
        fillColor: markerColor(item),
        fillOpacity: 0.9
      }).addTo(markersLayer).bindPopup(`<strong>${esc(item.name)}</strong><br>${esc(item.code || "Codice non disponibile")} · ${esc(item.municipality || "Comune non disponibile")}<br>${esc(formatNumber(item.designPopulation))} A.E.<br><button class="btn btn-primary" type="button" data-wastewater-plant-index="${index}">APRI SCHEDA</button>`);
      bounds.extend([item.lat, item.lon]);
    });
    if (bounds.isValid()) map.fitBounds(bounds.pad(0.08), { animate: false, maxZoom: 16 });
    else map.setView([44.55, 11.1], 8, { animate: false });
    mapStatus.textContent = `${items.length} ${items.length === 1 ? "depuratore visualizzato" : "depuratori visualizzati"}.`;
    resizeMap();
  }

  function renderResults(items) {
    const shown = items.slice(0, RESULT_LIST_LIMIT);
    resultsNode.innerHTML = shown.map((item, index) => `<article class="wastewater-plant-result">
      <div><small>${esc(item.code || "Senza codice")} · ${esc(item.provinceAbbreviation || item.province)}</small><h2>${esc(item.name)}</h2><p>${esc(item.municipality || "Comune non disponibile")} · ${esc(item.manager || "Gestore non disponibile")} · ${esc(formatNumber(item.designPopulation))} A.E.</p></div>
      <button class="btn" type="button" data-wastewater-result-index="${index}">MOSTRA E APRI SCHEDA</button>
    </article>`).join("");
    if (items.length > RESULT_LIST_LIMIT) {
      resultsNode.insertAdjacentHTML("beforeend", `<p class="wastewater-plant-result wastewater-plants-list-note">Nell’elenco sono mostrati i primi ${RESULT_LIST_LIMIT} risultati; tutti i ${items.length} depuratori sono presenti sulla mappa. Usa nome, codice, Comune o gestore per restringere la ricerca.</p>`);
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
    const message = [
      "🏭 *SCHEDA DEPURATORE*",
      "",
      `• *Impianto:* ${selectedItem.name}`,
      `• *Codice:* ${selectedItem.code || "Non disponibile"}`,
      `• *Comune:* ${selectedItem.municipality || "Non disponibile"}`,
      `• *Gestore:* ${selectedItem.manager || "Non disponibile"}`,
      `• *Tipo:* ${selectedItem.typeDescription || selectedItem.typeCode || "Non disponibile"}`,
      `• *Capacità:* ${formatNumber(selectedItem.designPopulation)} A.E.`,
      "",
      "📍 *NAVIGA VERSO IL DEPURATORE*",
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

  function runSearch() {
    const items = filterPlants();
    renderResults(items);
    showItems(items);
    if (!items.length) return setStatus("Nessun depuratore corrisponde ai filtri. Prova un nome più breve o seleziona tutta la Regione.", "error");
    setStatus(`${items.length} ${items.length === 1 ? "depuratore trovato" : "depuratori trovati"} su ${allPlants.length} impianti ARPAE.`, "success");
  }

  async function openPage() {
    $("menu-close-btn")?.click();
    $("home-page")?.classList.add("hidden");
    page?.classList.remove("hidden");
    page?.setAttribute("aria-hidden", "false");
    initializeMap();
    resizeMap();
    try {
      await loadAllPlants();
      runSearch();
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

  form?.addEventListener("submit", (event) => { event.preventDefault(); if (allPlants.length) runSearch(); else openPage(); });
  provinceInput?.addEventListener("change", () => { if (allPlants.length) runSearch(); });
  $("wastewater-plants-clear-btn")?.addEventListener("click", () => {
    queryInput.value = "";
    provinceInput.value = "";
    if (allPlants.length) runSearch();
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

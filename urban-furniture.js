(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const page = $("urban-furniture-page");
  const form = $("urban-furniture-form");
  const municipalityInput = $("urban-furniture-municipality");
  const categoryInput = $("urban-furniture-category");
  const statusNode = $("urban-furniture-status");
  const resultsNode = $("urban-furniture-results");
  const mapCard = $("urban-furniture-map-card");
  const mapNode = $("urban-furniture-map");
  const mapStatus = $("urban-furniture-map-status");
  const mapStyle = $("urban-furniture-map-style");
  const fullscreenButton = $("urban-furniture-fullscreen-btn");
  const sheet = $("urban-furniture-sheet");
  const sheetTitle = $("urban-furniture-sheet-title");
  const sheetBody = $("urban-furniture-sheet-body");
  const navigateButton = $("urban-furniture-navigate");
  const CACHE_PREFIX = "varga-urban-furniture:";
  const CACHE_TTL_MS = 10 * 60 * 1000;
  const OVERPASS_URL = "/api/urban-furniture";
  const MAX_RESULTS = 500;

  const CATEGORIES = Object.freeze({
    bench: { label: "Panchina", icon: "🪑", clause: '["amenity"="bench"]' },
    waste_basket: { label: "Cestino", icon: "🗑️", clause: '["amenity"="waste_basket"]' },
    drinking_water: { label: "Acqua potabile", icon: "🚰", clause: '["amenity"="drinking_water"]' },
    fountain: { label: "Fontana", icon: "⛲", clause: '["amenity"="fountain"]' },
    picnic_table: { label: "Tavolo da picnic", icon: "🧺", clause: '["leisure"="picnic_table"]' },
    bicycle_parking: { label: "Rastrelliera", icon: "🚲", clause: '["amenity"="bicycle_parking"]' },
    toilets: { label: "Bagno pubblico", icon: "🚻", clause: '["amenity"="toilets"]' },
    shelter: { label: "Pensilina o riparo", icon: "☂️", clause: '["amenity"="shelter"]' },
    street_lamp: { label: "Lampione", icon: "💡", clause: '["highway"="street_lamp"]' },
    playground: { label: "Area giochi", icon: "🛝", clause: '["leisure"="playground"]' },
    dog_park: { label: "Area cani", icon: "🐕", clause: '["leisure"="dog_park"]' },
    fitness_station: { label: "Area fitness", icon: "🏋️", clause: '["leisure"="fitness_station"]' },
    outdoor_seating: { label: "Sedute all’aperto", icon: "🪑", clause: '["leisure"="outdoor_seating"]' },
    bbq: { label: "Barbecue pubblico", icon: "♨️", clause: '["amenity"="bbq"]' },
    dog_toilet: { label: "Area igienica cani", icon: "🐾", clause: '["amenity"="dog_toilet"]' },
    recycling: { label: "Punto raccolta differenziata", icon: "♻️", clause: '["amenity"="recycling"]' },
    waste_disposal: { label: "Punto conferimento rifiuti", icon: "🚮", clause: '["amenity"="waste_disposal"]' },
    post_box: { label: "Cassetta postale", icon: "📮", clause: '["amenity"="post_box"]' },
    parcel_locker: { label: "Armadietto pacchi", icon: "📦", clause: '["amenity"="parcel_locker"]' },
    telephone: { label: "Telefono pubblico", icon: "☎️", clause: '["amenity"="telephone"]' },
    public_bookcase: { label: "Libreria pubblica", icon: "📚", clause: '["amenity"="public_bookcase"]' },
    bicycle_repair_station: { label: "Riparazione biciclette", icon: "🔧", clause: '["amenity"="bicycle_repair_station"]' },
    charging_station: { label: "Ricarica elettrica", icon: "🔌", clause: '["amenity"="charging_station"]' },
    motorcycle_parking: { label: "Parcheggio moto", icon: "🏍️", clause: '["amenity"="motorcycle_parking"]' },
    taxi: { label: "Area taxi", icon: "🚕", clause: '["amenity"="taxi"]' },
    compressed_air: { label: "Aria compressa", icon: "💨", clause: '["amenity"="compressed_air"]' },
    shower: { label: "Doccia pubblica", icon: "🚿", clause: '["amenity"="shower"]' },
    water_point: { label: "Punto acqua", icon: "💧", clause: '["amenity"="water_point"]' },
    clock: { label: "Orologio pubblico", icon: "🕐", clause: '["amenity"="clock"]' },
    grit_bin: { label: "Contenitore sale", icon: "🧂", clause: '["amenity"="grit_bin"]' },
    lounger: { label: "Sdraio pubblica", icon: "🏖️", clause: '["amenity"="lounger"]' },
    give_box: { label: "Box del riuso", icon: "🎁", clause: '["amenity"="give_box"]' },
    fire_hydrant: { label: "Idrante", icon: "🧯", clause: '["emergency"="fire_hydrant"]' },
    defibrillator: { label: "Defibrillatore", icon: "❤️", clause: '["emergency"="defibrillator"]' },
    bollard: { label: "Dissuasore", icon: "🚧", clause: '["barrier"="bollard"]' },
    cycle_barrier: { label: "Barriera ciclabile", icon: "🚲", clause: '["barrier"="cycle_barrier"]' },
    bus_stop: { label: "Fermata autobus", icon: "🚏", clause: '["highway"="bus_stop"]' },
    artwork: { label: "Opera d’arte urbana", icon: "🗿", clause: '["tourism"="artwork"]' },
    information: { label: "Punto informativo", icon: "ℹ️", clause: '["tourism"="information"]' }
  });

  const TAG_LABELS = Object.freeze({
    name: "Nome", "name:it": "Nome italiano", description: "Descrizione", operator: "Gestore", owner: "Proprietario",
    material: "Materiale", colour: "Colore", color: "Colore", condition: "Condizione", status: "Stato",
    access: "Accesso", fee: "A pagamento", wheelchair: "Accessibile in sedia a rotelle", covered: "Coperto",
    lit: "Illuminato", supervised: "Sorvegliato", opening_hours: "Orari di apertura", start_date: "Data di installazione",
    capacity: "Capienza", seats: "Posti a sedere", backrest: "Schienale", direction: "Orientamento",
    surface: "Superficie", level: "Piano", indoor: "Al coperto", drinking_water: "Acqua potabile",
    bottle: "Riempimento bottiglie", phone: "Telefono", email: "Email", website: "Sito web",
    ref: "Codice di riferimento", note: "Nota", fixme: "Dato da verificare", source: "Fonte dati",
    "addr:housenumber": "Numero civico", "addr:street": "Via", "addr:place": "Località", "addr:postcode": "CAP",
    "addr:city": "Comune", "addr:suburb": "Quartiere", "addr:province": "Provincia", "addr:country": "Paese",
    playground: "Tipo di gioco", recycling_type: "Tipo di raccolta", artwork_type: "Tipo di opera",
    artist_name: "Artista", inscription: "Iscrizione", tourism: "Categoria turistica", amenity: "Servizio",
    leisure: "Tempo libero", highway: "Elemento stradale", emergency: "Emergenza", barrier: "Barriera",
    shelter_type: "Tipo di riparo", bicycle_parking: "Tipo rastrelliera", socket: "Presa",
    "socket:type2": "Prese Tipo 2", "socket:chademo": "Prese CHAdeMO", "socket:ccs": "Prese CCS",
    payment: "Pagamento", "payment:cash": "Pagamento in contanti", "payment:contactless": "Pagamento contactless",
    surveillance: "Videosorveglianza", camera: "Tipo di telecamera", "fire_hydrant:type": "Tipo di idrante"
  });

  const YES_NO_VALUES = Object.freeze({ yes: "Sì", no: "No", limited: "Limitato", permissive: "Consentito", private: "Privato", customers: "Solo clienti" });

  const TILE_LAYERS = {
    classic: { url: "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "&copy; OpenStreetMap contributors" } },
    satellite: { url: "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "Tiles &copy; Esri" } },
    labels: { url: "https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "Labels &copy; Esri", pane: "overlayPane" } }
  };

  let map = null;
  let baseLayer = null;
  let hybridLabels = null;
  let markersLayer = null;
  let operatorMarker = null;
  let fullscreen = false;
  let currentItems = [];
  let selectedItem = null;

  const esc = (value) => String(value ?? "—").replace(/[&<>'"]/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" }[char]));
  const setStatus = (message, type = "") => {
    statusNode.textContent = message;
    statusNode.className = `urban-furniture-status ${type}`.trim();
  };

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
    map = L.map(mapNode, { zoomControl: true, zoomAnimation: false, fadeAnimation: false, markerZoomAnimation: false }).setView([44.4949, 11.3426], 13);
    markersLayer = L.layerGroup().addTo(map);
    applyMapStyle(mapStyle?.value || "classic");
  }

  function resizeMap() {
    requestAnimationFrame(() => map?.invalidateSize({ pan: false, animate: false }));
    setTimeout(() => map?.invalidateSize({ pan: false, animate: false }), 180);
  }

  function setFullscreen(active) {
    fullscreen = Boolean(active);
    mapCard?.classList.toggle("urban-furniture-map-card--fullscreen", fullscreen);
    document.body.classList.toggle("urban-furniture-fullscreen-open", fullscreen);
    fullscreenButton?.setAttribute("aria-pressed", String(fullscreen));
    if (fullscreenButton) fullscreenButton.textContent = fullscreen ? "✕ CHIUDI MAPPA" : "⛶ SCHERMO INTERO";
    resizeMap();
  }

  function openPage() {
    $("menu-close-btn")?.click();
    $("home-page")?.classList.add("hidden");
    page?.classList.remove("hidden");
    page?.setAttribute("aria-hidden", "false");
    initializeMap();
    resizeMap();
    municipalityInput?.focus();
  }

  function closePage() {
    setFullscreen(false);
    closeSheet();
    page?.classList.add("hidden");
    page?.setAttribute("aria-hidden", "true");
    $("home-page")?.classList.remove("hidden");
  }

  function categoryFor(tags = {}) {
    return Object.entries(CATEGORIES).find(([, data]) => {
      const match = data.clause.match(/\["([^"]+)"="([^"]+)"\]/);
      return match && tags[match[1]] === match[2];
    })?.[0] || "bench";
  }

  function normalizeElement(element) {
    const lat = Number(element.lat ?? element.center?.lat);
    const lon = Number(element.lon ?? element.center?.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    const tags = element.tags || {};
    const category = categoryFor(tags);
    const details = CATEGORIES[category];
    return {
      id: `${element.type}/${element.id}`,
      osmId: element.id,
      osmType: element.type,
      lat,
      lon,
      category,
      label: details.label,
      icon: details.icon,
      name: tags.name || tags["name:it"] || details.label,
      street: tags["addr:street"] || tags.location || tags.description || "",
      tags
    };
  }

  function cacheKey(query) { return `${CACHE_PREFIX}${query}`; }
  function readCache(query) {
    try {
      const cached = JSON.parse(sessionStorage.getItem(cacheKey(query)) || "null");
      return cached && Date.now() - cached.savedAt < CACHE_TTL_MS ? cached.items : null;
    } catch (_) { return null; }
  }
  function writeCache(query, items) {
    try { sessionStorage.setItem(cacheKey(query), JSON.stringify({ savedAt: Date.now(), items })); } catch (_) {}
  }

  async function requestOverpass(params, cacheId) {
    const cached = readCache(cacheId);
    if (cached) return cached;
    const response = await fetch(`${OVERPASS_URL}?${new URLSearchParams(params).toString()}`, { headers: { Accept: "application/json" } });
    const payload = await response.json().catch(() => null);
    if (!response.ok || payload?.ok === false) throw new Error(payload?.error || `Servizio cartografico momentaneamente non disponibile (${response.status}).`);
    const seen = new Set();
    const items = (payload.elements || []).map(normalizeElement).filter(Boolean).filter((item) => {
      if (seen.has(item.id)) return false;
      seen.add(item.id);
      return true;
    }).slice(0, MAX_RESULTS);
    writeCache(cacheId, items);
    return items;
  }

  async function searchMunicipality(municipality, category) {
    return requestOverpass({ mode: "municipality", municipality, category }, `municipality:${municipality.toLocaleLowerCase("it-IT")}:${category}`);
  }

  async function searchNearby(lat, lon, category) {
    return requestOverpass({ mode: "nearby", lat: lat.toFixed(6), lon: lon.toFixed(6), category }, `nearby:${lat.toFixed(3)}:${lon.toFixed(3)}:${category}`);
  }

  function markerIcon(item) {
    return L.divIcon({ className: "", html: `<span class="urban-furniture-marker" title="${esc(item.label)}">${item.icon}</span>`, iconSize: [34, 34], iconAnchor: [17, 17] });
  }

  function showItems(items, center = null) {
    initializeMap();
    markersLayer.clearLayers();
    currentItems = items;
    const bounds = L.latLngBounds([]);
    items.forEach((item, index) => {
      L.marker([item.lat, item.lon], { icon: markerIcon(item) }).addTo(markersLayer)
        .bindPopup(`<strong>${item.icon} ${esc(item.name)}</strong><br>${esc(item.label)}${item.street ? `<br>${esc(item.street)}` : ""}<br><button class="btn btn-primary" type="button" data-urban-furniture-index="${index}">APRI SCHEDA</button>`);
      bounds.extend([item.lat, item.lon]);
    });
    if (center) {
      operatorMarker?.remove();
      operatorMarker = L.circleMarker(center, { radius: 8, color: "#fff", weight: 3, fillColor: "#1769e0", fillOpacity: 1 }).addTo(map).bindPopup("La mia posizione");
      bounds.extend(center);
    }
    if (bounds.isValid()) map.fitBounds(bounds.pad(0.08), { animate: false, maxZoom: 17 });
    mapStatus.textContent = `${items.length} ${items.length === 1 ? "elemento visualizzato" : "elementi visualizzati"}.`;
    resizeMap();
  }

  function renderResults(items) {
    resultsNode.innerHTML = items.slice(0, 60).map((item, index) => `<article class="urban-furniture-result"><h2>${item.icon} ${esc(item.name)}</h2><p>${esc(item.label)}${item.street ? ` · ${esc(item.street)}` : ""}</p><button class="btn" type="button" data-urban-result-index="${index}">MOSTRA E APRI SCHEDA</button></article>`).join("");
    if (items.length > 60) resultsNode.insertAdjacentHTML("beforeend", `<p class="urban-furniture-result">Sono mostrati i primi 60 risultati nell’elenco; tutti i ${items.length} elementi sono presenti sulla mappa.</p>`);
    resultsNode.classList.remove("hidden");
  }

  function humanizeTag(key) {
    return TAG_LABELS[key] || key.replaceAll(":", " · ").replaceAll("_", " ").replace(/\b\w/g, (letter) => letter.toLocaleUpperCase("it-IT"));
  }

  function humanizeValue(value) {
    const text = String(value ?? "").trim();
    return text.split(";").map((part) => YES_NO_VALUES[part.trim().toLocaleLowerCase("en-US")] || part.trim().replaceAll("_", " ")).filter(Boolean).join(", ") || "—";
  }

  function detailRow(label, value, options = {}) {
    if (value === undefined || value === null || String(value).trim() === "") return "";
    const content = options.raw ? String(value) : esc(value);
    return `<div class="urban-furniture-detail-row"><dt>${esc(label)}</dt><dd>${content}</dd></div>`;
  }

  function renderTagDetails(item) {
    const tags = item.tags || {};
    const address = [tags["addr:street"], tags["addr:housenumber"], tags["addr:postcode"], tags["addr:city"]].filter(Boolean).join(" ");
    const featuredKeys = ["description", "operator", "owner", "material", "colour", "color", "condition", "status", "access", "fee", "wheelchair", "opening_hours", "capacity", "seats", "backrest", "covered", "lit", "supervised", "surface", "start_date", "phone", "email", "website", "ref", "note"];
    const hiddenKeys = new Set(["name", "name:it", "addr:street", "addr:housenumber", "addr:postcode", "addr:city", ...featuredKeys]);
    const featured = featuredKeys.map((key) => tags[key] === undefined ? "" : detailRow(humanizeTag(key), humanizeValue(tags[key]))).join("");
    const remaining = Object.entries(tags).filter(([key, value]) => !hiddenKeys.has(key) && value !== "").sort(([a], [b]) => humanizeTag(a).localeCompare(humanizeTag(b), "it")).map(([key, value]) => detailRow(humanizeTag(key), humanizeValue(value))).join("");
    const osmUrl = `https://www.openstreetmap.org/${encodeURIComponent(item.osmType)}/${encodeURIComponent(item.osmId)}`;
    return [
      `<section class="urban-furniture-detail-section"><h3>Informazioni principali</h3><dl>${detailRow("Tipo", item.label)}${detailRow("Nome", item.name)}${detailRow("Comune cercato", municipalityInput.value.trim() || "Ricerca vicino a me")}${detailRow("Indirizzo", address || item.street)}${detailRow("Coordinate", `${item.lat.toFixed(6)}, ${item.lon.toFixed(6)}`)}</dl></section>`,
      featured ? `<section class="urban-furniture-detail-section"><h3>Caratteristiche e servizi</h3><dl>${featured}</dl></section>` : "",
      remaining ? `<section class="urban-furniture-detail-section"><h3>Tutti i dati disponibili</h3><dl>${remaining}</dl></section>` : "",
      `<section class="urban-furniture-detail-section"><h3>Dati cartografici</h3><dl>${detailRow("Riferimento OSM", item.id)}${detailRow("Tipo geometria", humanizeValue(item.osmType))}${detailRow("Scheda OpenStreetMap", `<a href="${osmUrl}" target="_blank" rel="noopener">Apri la fonte originale</a>`, { raw: true })}</dl></section>`
    ].join("");
  }

  function openSheet(item) {
    if (!item) return;
    selectedItem = item;
    sheetTitle.textContent = `${item.icon} ${item.name}`;
    sheetBody.innerHTML = renderTagDetails(item);
    navigateButton.href = `https://www.google.com/maps/dir/?api=1&destination=${item.lat},${item.lon}`;
    sheet.classList.remove("hidden");
    sheet.setAttribute("aria-hidden", "false");
  }

  function closeSheet() {
    selectedItem = null;
    sheet?.classList.add("hidden");
    sheet?.setAttribute("aria-hidden", "true");
  }

  async function openStreetView() {
    if (!selectedItem) return;
    const api = window.HeraStreetViewCards;
    if (typeof api?.openForCoordinates !== "function") return window.alert("Vista 360° non disponibile in questo momento. Riprova tra qualche secondo.");
    await api.openForCoordinates({ lat: selectedItem.lat, lng: selectedItem.lon }, $("urban-furniture-street-view"), {
      targetLabel: selectedItem.label,
      modalTitle: `🌐 Vista 360° ${selectedItem.label.toLocaleLowerCase("it-IT")}`
    });
  }

  async function shareSelectedItem() {
    if (!selectedItem) return;
    const navigationUrl = `https://www.google.com/maps/dir/?api=1&destination=${selectedItem.lat},${selectedItem.lon}`;
    const message = [`🪑 *ARREDO URBANO*`, "", `• *Elemento:* ${selectedItem.name}`, `• *Tipo:* ${selectedItem.label}`, `• *Coordinate:* ${selectedItem.lat.toFixed(6)}, ${selectedItem.lon.toFixed(6)}`, "", "📍 *NAVIGA VERSO L’ELEMENTO*", navigationUrl].join("\n");
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

  async function executeMunicipalitySearch() {
    const municipality = municipalityInput.value.trim();
    const category = categoryInput.value;
    if (municipality.length < 2) return setStatus("Inserisci il nome completo del Comune.", "error");
    setStatus("Cerco l’arredo urbano nel Comune…");
    resultsNode.classList.add("hidden");
    try {
      const items = await searchMunicipality(municipality, category);
      if (!items.length) throw new Error("Nessun elemento censito. Prova un’altra categoria o verifica il nome del Comune.");
      renderResults(items);
      showItems(items);
      setStatus(`${items.length} ${items.length === 1 ? "elemento trovato" : "elementi trovati"}.`, "success");
    } catch (error) { setStatus(error.message || "Ricerca non riuscita.", "error"); }
  }

  function locateAndSearch() {
    if (!navigator.geolocation) return setStatus("La posizione non è disponibile su questo dispositivo.", "error");
    setStatus("Cerco la posizione dell’operatore…");
    navigator.geolocation.getCurrentPosition(async ({ coords }) => {
      const center = [Number(coords.latitude), Number(coords.longitude)];
      setStatus("Cerco gli elementi entro 3 km dalla tua posizione…");
      try {
        const items = await searchNearby(center[0], center[1], categoryInput.value);
        if (!items.length) throw new Error("Nessun elemento censito entro 3 km dalla tua posizione.");
        renderResults(items);
        showItems(items, center);
        setStatus(`${items.length} elementi trovati entro 3 km.`, "success");
      } catch (error) { setStatus(error.message || "Ricerca non riuscita.", "error"); }
    }, (error) => setStatus(error.code === 1 ? "Permesso posizione negato. Abilitalo nelle impostazioni." : "Non riesco a rilevare la posizione.", "error"), { enableHighAccuracy: true, timeout: 12000, maximumAge: 60000 });
  }

  form?.addEventListener("submit", (event) => { event.preventDefault(); executeMunicipalitySearch(); });
  $("urban-furniture-nearby-btn")?.addEventListener("click", locateAndSearch);
  $("open-urban-furniture-btn")?.addEventListener("click", openPage);
  $("urban-furniture-back-btn")?.addEventListener("click", closePage);
  fullscreenButton?.addEventListener("click", () => setFullscreen(!fullscreen));
  mapStyle?.addEventListener("change", () => applyMapStyle(mapStyle.value));
  resultsNode?.addEventListener("click", (event) => {
    const button = event.target.closest?.("[data-urban-result-index]");
    if (!button) return;
    const item = currentItems[Number(button.dataset.urbanResultIndex)];
    map?.setView([item.lat, item.lon], 18, { animate: false });
    openSheet(item);
  });
  mapNode?.addEventListener("click", (event) => {
    const button = event.target.closest?.("[data-urban-furniture-index]");
    if (button) openSheet(currentItems[Number(button.dataset.urbanFurnitureIndex)]);
  });
  $("urban-furniture-sheet-close")?.addEventListener("click", closeSheet);
  $("urban-furniture-street-view")?.addEventListener("click", openStreetView);
  $("urban-furniture-whazzup")?.addEventListener("click", shareSelectedItem);
  sheet?.addEventListener("click", (event) => { if (event.target === sheet) closeSheet(); });
  document.addEventListener("keydown", (event) => { if (event.key === "Escape") fullscreen ? setFullscreen(false) : closeSheet(); });
})();

(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const page = $("green-areas-page");
  const form = $("green-areas-form");
  const nameInput = $("green-areas-name");
  const municipalityInput = $("green-areas-municipality");
  const status = $("green-areas-status");
  const resultsNode = $("green-areas-results");
  const mapCard = $("green-areas-map-card");
  const mapNode = $("green-areas-map");
  const mapStatus = $("green-areas-map-status");
  const mapStyle = $("green-areas-map-style");
  const fullscreenButton = $("green-areas-fullscreen-btn");
  const REGION_VIEWBOX = "9.1729,45.1360,12.7556,43.7310";
  const CACHE_PREFIX = "varga-green-area-search:";
  const CACHE_TTL_MS = 10 * 60 * 1000;
  const OVERPASS_URL = "https://overpass-api.de/api/interpreter";
  const GREEN_TYPES = new Set(["park", "garden", "nature_reserve", "recreation_ground", "grass", "village_green", "forest", "wood", "playground"]);
  let map = null;
  let baseLayer = null;
  let hybridLabels = null;
  let officialGreenLayer = null;
  let searchLayer = null;
  let fullscreen = false;

  const esc = (value) => String(value ?? "—").replace(/[&<>'"]/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[char]));
  const setStatus = (message, type = "") => {
    status.textContent = message;
    status.className = `green-areas-status ${type}`.trim();
  };

  const TILE_LAYERS = {
    classic: { url: "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "&copy; OpenStreetMap" } },
    satellite: { url: "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "Tiles &copy; Esri" } },
    labels: { url: "https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "Labels &copy; Esri", pane: "overlayPane" } }
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
    if (map || !window.L) return;
    map = L.map(mapNode, { zoomControl: true, zoomAnimation: false, fadeAnimation: false, markerZoomAnimation: false }).setView([44.55, 11.05], 8);
    searchLayer = L.layerGroup().addTo(map);
    applyMapStyle(mapStyle?.value || "classic");
    officialGreenLayer = L.tileLayer.wms("https://servizigis.regione.emilia-romagna.it/wms/dbtr", {
      layers: "PSR_Area_verde",
      format: "image/png",
      transparent: true,
      version: "1.3.0",
      opacity: 0.58,
      maxZoom: 20,
      attribution: "Aree verdi: DBTR Regione Emilia-Romagna (CC BY 4.0)"
    }).addTo(map);
    officialGreenLayer.on("tileerror", () => {
      mapStatus.textContent = "Il livello ufficiale regionale non risponde. Ricerca e mappa di base restano disponibili.";
    });
  }

  function resizeMap() {
    requestAnimationFrame(() => map?.invalidateSize({ pan: false, animate: false }));
    setTimeout(() => map?.invalidateSize({ pan: false, animate: false }), 180);
  }

  function setFullscreen(active) {
    fullscreen = Boolean(active);
    mapCard?.classList.toggle("green-areas-map-card--fullscreen", fullscreen);
    document.body.classList.toggle("green-areas-fullscreen-open", fullscreen);
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
    nameInput?.focus();
  }

  function closePage() {
    setFullscreen(false);
    page?.classList.add("hidden");
    page?.setAttribute("aria-hidden", "true");
    $("home-page")?.classList.remove("hidden");
  }

  function cacheKey(query) {
    return `${CACHE_PREFIX}${query.toLocaleLowerCase("it-IT")}`;
  }

  function readCache(query) {
    try {
      const cached = JSON.parse(sessionStorage.getItem(cacheKey(query)) || "null");
      return cached && Date.now() - cached.savedAt < CACHE_TTL_MS ? cached.results : null;
    } catch (_) { return null; }
  }

  function writeCache(query, results) {
    try { sessionStorage.setItem(cacheKey(query), JSON.stringify({ savedAt: Date.now(), results })); } catch (_) {}
  }

  function isGreenArea(item) {
    const type = String(item.type || item.addresstype || "").toLowerCase();
    return GREEN_TYPES.has(type) || /parco|giardino|area verde|bosco|riserva/i.test(`${item.name || ""} ${item.display_name || ""}`);
  }

  function escapeOverpassString(value) {
    return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/"/g, '\\"');
  }

  function normalizeOverpassItem(element) {
    const tags = element.tags || {};
    const lat = Number(element.lat ?? element.center?.lat);
    const lon = Number(element.lon ?? element.center?.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
    const box = element.bounds;
    const street = tags["addr:street"] || tags.loc_name || "";
    const title = tags.name || tags["name:it"] || (street ? `Area verde – ${street}` : "Area verde senza nome");
    return {
      osm_type: element.type,
      osm_id: element.id,
      name: title,
      lat: String(lat),
      lon: String(lon),
      boundingbox: box ? [box.minlat, box.maxlat, box.minlon, box.maxlon].map(String) : null,
      display_name: [title, street, tags["addr:city"]].filter((value, index, values) => value && values.indexOf(value) === index).join(", "),
      address: { city: tags["addr:city"] || "" },
      tags,
      source: "OpenStreetMap"
    };
  }

  async function searchMunicipalGreenAreas(name, municipality) {
    const query = `overpass:${municipality}:${name}`;
    const cached = readCache(query);
    if (cached) return cached;
    const town = escapeOverpassString(municipality);
    const overpassQuery = `[out:json][timeout:25];area["ISO3166-2"="IT-45"]["boundary"="administrative"]->.region;area(area.region)["boundary"="administrative"]["admin_level"="8"]["name"~"^${town}$","i"]->.municipality;(nwr["leisure"~"^(park|garden|recreation_ground|playground|dog_park)$"](area.municipality);nwr["landuse"~"^(grass|village_green|recreation_ground|forest|allotments)$"](area.municipality);nwr["natural"="wood"](area.municipality);nwr["boundary"="protected_area"](area.municipality););out center tags bb;`;
    const response = await fetch(OVERPASS_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8", Accept: "application/json" },
      body: new URLSearchParams({ data: overpassQuery })
    });
    if (!response.ok) throw new Error(`Catasto cartografico non disponibile (${response.status}).`);
    const payload = await response.json();
    const needle = name.toLocaleLowerCase("it-IT");
    const seen = new Set();
    const items = (payload.elements || []).map(normalizeOverpassItem).filter(Boolean).filter((item) => {
      const key = `${item.osm_type}:${item.osm_id}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return !needle || `${item.name} ${item.display_name}`.toLocaleLowerCase("it-IT").includes(needle);
    });
    writeCache(query, items);
    return items;
  }

  async function searchGreenAreas(name, municipality) {
    const query = [name, municipality, "Emilia-Romagna", "Italia"].filter(Boolean).join(", ");
    const cached = readCache(query);
    if (cached) return cached;
    const params = new URLSearchParams({
      format: "jsonv2",
      q: query,
      limit: "20",
      countrycodes: "it",
      viewbox: REGION_VIEWBOX,
      bounded: "1",
      addressdetails: "1",
      "accept-language": "it"
    });
    const response = await fetch(`https://nominatim.openstreetmap.org/search?${params}`, { headers: { Accept: "application/json" } });
    if (!response.ok) throw new Error(`Servizio di ricerca non disponibile (${response.status}).`);
    const payload = await response.json();
    const filtered = payload.filter(isGreenArea);
    writeCache(query, filtered);
    return filtered;
  }

  function boundsFromResult(item) {
    const box = item.boundingbox?.map(Number);
    return box?.length === 4 && box.every(Number.isFinite) ? L.latLngBounds([box[0], box[2]], [box[1], box[3]]) : null;
  }

  function showArea(item) {
    initializeMap();
    const lat = Number(item.lat);
    const lon = Number(item.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
    searchLayer.clearLayers();
    const bounds = boundsFromResult(item);
    if (bounds) L.rectangle(bounds, { color: "#08783f", weight: 3, fillColor: "#31b96b", fillOpacity: 0.16 }).addTo(searchLayer);
    L.marker([lat, lon]).addTo(searchLayer).bindPopup(`<strong>${esc(item.name || item.display_name.split(",")[0])}</strong><br>${esc(item.display_name)}<br><a href="https://www.google.com/maps/dir/?api=1&destination=${lat},${lon}" target="_blank" rel="noopener">NAVIGA</a>`).openPopup();
    if (bounds?.isValid()) map.fitBounds(bounds.pad(0.12), { animate: false, maxZoom: 17 });
    else map.setView([lat, lon], 17, { animate: false });
    mapStatus.textContent = `${item.name || item.display_name.split(",")[0]} evidenziata sulla mappa.`;
    resizeMap();
  }

  function showAllAreas(items, municipality) {
    initializeMap();
    searchLayer.clearLayers();
    const combinedBounds = L.latLngBounds([]);
    items.forEach((item) => {
      const lat = Number(item.lat);
      const lon = Number(item.lon);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
      const title = item.name || item.display_name.split(",")[0];
      const bounds = boundsFromResult(item);
      if (bounds?.isValid()) {
        L.rectangle(bounds, { color: "#08783f", weight: 1.5, fillColor: "#31b96b", fillOpacity: 0.13 }).addTo(searchLayer)
          .bindPopup(`<strong>${esc(title)}</strong><br>${esc(item.display_name)}<br><a href="https://www.google.com/maps/dir/?api=1&destination=${lat},${lon}" target="_blank" rel="noopener">NAVIGA</a>`);
        combinedBounds.extend(bounds);
      } else {
        L.circleMarker([lat, lon], { radius: 6, color: "#08783f", weight: 2, fillColor: "#31b96b", fillOpacity: 0.72 }).addTo(searchLayer)
          .bindPopup(`<strong>${esc(title)}</strong><br>${esc(item.display_name)}<br><a href="https://www.google.com/maps/dir/?api=1&destination=${lat},${lon}" target="_blank" rel="noopener">NAVIGA</a>`);
        combinedBounds.extend([lat, lon]);
      }
    });
    if (combinedBounds.isValid()) map.fitBounds(combinedBounds.pad(0.08), { animate: false, maxZoom: 15 });
    mapStatus.textContent = `${items.length} aree verdi di ${municipality} evidenziate insieme sulla mappa.`;
    resizeMap();
  }

  function renderResults(items) {
    resultsNode.innerHTML = items.map((item, index) => {
      const title = item.name || item.display_name.split(",")[0];
      const municipality = item.address?.city || item.address?.town || item.address?.village || item.address?.municipality || municipalityInput.value.trim() || "Comune non indicato";
      return `<article class="green-area-result"><div><small>${esc(municipality)}</small><h2>${esc(title)}</h2><p>${esc(item.display_name)}</p></div><button class="btn btn-primary" type="button" data-green-area-index="${index}">MOSTRA SULLA MAPPA</button></article>`;
    }).join("");
    resultsNode.classList.remove("hidden");
    resultsNode.querySelectorAll("[data-green-area-index]").forEach((button) => {
      button.addEventListener("click", () => showArea(items[Number(button.dataset.greenAreaIndex)]));
    });
  }

  form?.addEventListener("submit", async (event) => {
    event.preventDefault();
    const name = nameInput.value.trim();
    const municipality = municipalityInput.value.trim();
    if (!municipality && name.length < 2) return setStatus("Indica il Comune oppure inserisci almeno 2 caratteri del nome.", "error");
    if (name && name.length < 2) return setStatus("Inserisci almeno 2 caratteri del nome oppure lascia il campo vuoto.", "error");
    setStatus("Ricerca nelle aree verdi dell’Emilia-Romagna…");
    resultsNode.classList.add("hidden");
    try {
      const items = municipality ? await searchMunicipalGreenAreas(name, municipality) : await searchGreenAreas(name, municipality);
      if (!items.length) throw new Error(name ? "Nessuna area verde con questo nome. Prova con un nome più breve o lascia vuoto il nome per vedere tutto il Comune." : "Nessuna area verde cartografata per questo Comune.");
      renderResults(items);
      if (municipality && !name) showAllAreas(items, municipality);
      else showArea(items[0]);
      setStatus(`${items.length} ${items.length === 1 ? "area verde trovata" : "aree verdi trovate"}.`, "success");
    } catch (error) {
      setStatus(error.message || "Ricerca non riuscita.", "error");
    }
  });

  $("open-green-areas-btn")?.addEventListener("click", openPage);
  $("green-areas-back-btn")?.addEventListener("click", closePage);
  mapStyle?.addEventListener("change", () => applyMapStyle(mapStyle.value));
  fullscreenButton?.addEventListener("click", () => setFullscreen(!fullscreen));
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && fullscreen) setFullscreen(false);
  });
})();

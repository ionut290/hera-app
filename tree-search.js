(() => {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const page = $("tree-search-page");
  const form = $("tree-search-form");
  const municipality = $("tree-municipality");
  const number = $("tree-number");
  const status = $("tree-search-status");
  const result = $("tree-result");
  const mapCard = $("tree-map-card");
  const mapNode = $("tree-map");
  const mapStatus = $("tree-map-status");
  const mapStyle = $("tree-map-style");
  const mapLocationButton = $("tree-map-location-btn");
  const mapFullscreenButton = $("tree-map-fullscreen-btn");
  const dialog = $("tree-qr-dialog");
  const video = $("tree-qr-video");
  const qrStatus = $("tree-qr-status");
  let map = null;
  let marker = null;
  let treesLayer = null;
  let baseLayer = null;
  let hybridLabels = null;
  let userLocationMarker = null;
  let userAccuracyCircle = null;
  let viewportTimer = 0;
  let viewportRequest = 0;
  let viewportAbort = null;
  let lastViewportKey = "";
  let stream = null;
  let scanFrame = 0;
  let mapFullscreen = false;

  const esc = (value) => String(value ?? "—").replace(/[&<>'"]/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[char]));
  const setStatus = (message, type = "") => { status.textContent = message; status.className = `tree-search-status ${type}`.trim(); };
  const hideHome = (hidden) => $("home-page")?.classList.toggle("hidden", hidden);
  const normalizeTreeIdentifier = (value) => String(value ?? "").trim().replace(/\s+/g, "").toUpperCase();

  const TREE_FIELD_LABELS = {
    num_pt: "Numero punto",
    cod_alb: "Codice albero",
    classe: "Specie / classe",
    nome_scientifico: "Nome scientifico",
    nome_comune: "Nome comune",
    cl_h: "Classe di altezza",
    classe_circonferenza_diametro: "Classe circonferenza / diametro",
    circonferenza: "Circonferenza",
    diametro: "Diametro",
    altezza: "Altezza",
    quartiere: "Quartiere",
    via: "Via",
    indirizzo: "Indirizzo",
    localizzazione: "Localizzazione",
    dimora: "Dimora",
    pregio: "Albero di pregio",
    irrigazione: "Irrigazione",
    distanza_fabbricati: "Distanza dai fabbricati",
    data_impianto: "Data impianto",
    data_inventario: "Data inventario",
    data_aggiornamento: "Data aggiornamento",
    stato: "Stato",
    note: "Note",
    geo_point_2d: "Coordinate"
  };

  const TREE_FIELD_PRIORITY = [
    "num_pt", "cod_alb", "classe", "nome_scientifico", "nome_comune",
    "cl_h", "classe_circonferenza_diametro", "circonferenza", "diametro", "altezza",
    "quartiere", "via", "indirizzo", "localizzazione", "dimora", "pregio", "irrigazione",
    "distanza_fabbricati", "data_impianto", "data_inventario", "data_aggiornamento", "stato", "note", "geo_point_2d"
  ];

  function hasTreeValue(value) {
    if (value === null || value === undefined || value === "") return false;
    if (Array.isArray(value)) return value.length > 0;
    if (typeof value === "object") return Object.keys(value).length > 0;
    return true;
  }

  function treeFieldLabel(key) {
    if (TREE_FIELD_LABELS[key]) return TREE_FIELD_LABELS[key];
    return String(key || "")
      .replace(/^geo_/, "")
      .replace(/_/g, " ")
      .replace(/\b\w/g, (char) => char.toUpperCase());
  }

  function formatTreeValue(key, value) {
    if (key === "pregio") {
      const normalized = String(value).trim().toUpperCase();
      if (["S", "SI", "SÌ", "YES", "TRUE", "1"].includes(normalized)) return "Sì";
      if (["N", "NO", "FALSE", "0"].includes(normalized)) return "No";
    }
    if (key === "geo_point_2d" && value && typeof value === "object") {
      const lat = Number(value.lat);
      const lon = Number(value.lon);
      if (Number.isFinite(lat) && Number.isFinite(lon)) return `${lat.toFixed(6)}, ${lon.toFixed(6)}`;
    }
    if (Array.isArray(value)) return value.map((item) => typeof item === "object" ? JSON.stringify(item) : String(item)).join(", ");
    if (typeof value === "object") {
      try { return JSON.stringify(value); } catch (_) { return String(value); }
    }
    return String(value);
  }

  function buildTreeDetails(tree) {
    const keys = Object.keys(tree || {}).filter((key) => hasTreeValue(tree[key]));
    const ordered = [
      ...TREE_FIELD_PRIORITY.filter((key) => keys.includes(key)),
      ...keys.filter((key) => !TREE_FIELD_PRIORITY.includes(key)).sort((a, b) => treeFieldLabel(a).localeCompare(treeFieldLabel(b), "it"))
    ];
    return ordered.map((key) => `<div><span>${esc(treeFieldLabel(key))}</span><strong>${esc(formatTreeValue(key, tree[key]))}</strong></div>`).join("");
  }

  function resizeMap() {
    requestAnimationFrame(() => map?.invalidateSize({ pan: false, animate: false }));
    setTimeout(() => map?.invalidateSize({ pan: false, animate: false }), 180);
  }

  function setMapFullscreen(active) {
    if (!mapCard || !mapFullscreenButton) return;
    mapFullscreen = Boolean(active);
    mapCard.classList.toggle("tree-map-card--fullscreen", mapFullscreen);
    document.body.classList.toggle("tree-map-fullscreen-open", mapFullscreen);
    mapFullscreenButton.setAttribute("aria-pressed", String(mapFullscreen));
    mapFullscreenButton.textContent = mapFullscreen ? "✕ CHIUDI MAPPA" : "⛶ SCHERMO INTERO";
    resizeMap();
  }

  function openPage() {
    document.getElementById("menu-close-btn")?.click();
    hideHome(true);
    page.classList.remove("hidden");
    page.setAttribute("aria-hidden", "false");
    initializeMap();
    setTimeout(() => { map?.invalidateSize(); loadVisibleTrees(); }, 80);
    number.focus();
  }
  function closePage() {
    stopScanner();
    setMapFullscreen(false);
    page.classList.add("hidden");
    page.setAttribute("aria-hidden", "true");
    hideHome(false);
  }

  function parseQr(raw) {
    const text = String(raw || "").trim();
    if (!text) return null;
    try {
      const parsed = JSON.parse(text);
      const comune = parsed.comune || parsed.municipality || parsed.city;
      const numero = parsed.numero || parsed.numeroAlbero || parsed.treeId || parsed.id;
      if (comune && numero) return { comune, numero: String(numero) };
    } catch (_) {}
    try {
      const url = new URL(text);
      const comune = url.searchParams.get("comune") || url.searchParams.get("municipality");
      const numero = url.searchParams.get("numero") || url.searchParams.get("treeId") || url.searchParams.get("id");
      if (comune && numero) return { comune, numero };
    } catch (_) {}
    const match = text.match(/(?:comune\s*[:=]\s*)?([^|;,\n]+)[|;,\n]\s*(?:numero|albero|id)?\s*[:=#]?\s*([\w-]+)/i) || text.match(/^([A-Za-zÀ-ÿ '’-]+)[:/#-]([\w-]+)$/);
    return match ? { comune: match[1].trim(), numero: match[2].trim() } : null;
  }

  function acceptQr(raw) {
    const data = parseQr(raw);
    if (!data) { qrStatus.textContent = "QR non riconosciuto. Deve contenere Comune e numero dell’albero."; return; }
    const option = [...municipality.options].find((item) => item.value.toLowerCase() === data.comune.toLowerCase() && !item.disabled);
    if (!option) { qrStatus.textContent = `Il Comune “${data.comune}” non è ancora collegato.`; return; }
    municipality.value = option.value;
    number.value = data.numero;
    stopScanner();
    dialog.close();
    form.requestSubmit();
  }

  async function detector() {
    if (!("BarcodeDetector" in window)) throw new Error("Lettura QR non supportata da questo browser. Usa “Leggi QR da immagine” oppure inserisci il numero.");
    const formats = await BarcodeDetector.getSupportedFormats();
    if (!formats.includes("qr_code")) throw new Error("Questo dispositivo non supporta la lettura QR.");
    return new BarcodeDetector({ formats: ["qr_code"] });
  }

  async function startScanner() {
    dialog.showModal();
    qrStatus.textContent = "Avvio fotocamera…";
    try {
      const reader = await detector();
      stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: "environment" } }, audio: false });
      video.srcObject = stream;
      await video.play();
      qrStatus.textContent = "Inquadra il QR applicato all’albero.";
      const scan = async () => {
        if (!stream || dialog.open === false) return;
        try {
          const codes = await reader.detect(video);
          if (codes[0]?.rawValue) return acceptQr(codes[0].rawValue);
        } catch (_) {}
        scanFrame = requestAnimationFrame(scan);
      };
      scan();
    } catch (error) { qrStatus.textContent = error.message || "Impossibile aprire la fotocamera."; }
  }
  function stopScanner() {
    cancelAnimationFrame(scanFrame);
    stream?.getTracks().forEach((track) => track.stop());
    stream = null;
    if (video) video.srcObject = null;
  }

  async function scanFile(file) {
    try {
      const reader = await detector();
      const bitmap = await createImageBitmap(file);
      const codes = await reader.detect(bitmap);
      bitmap.close();
      if (!codes[0]?.rawValue) throw new Error("Nessun QR leggibile nell’immagine.");
      acceptQr(codes[0].rawValue);
    } catch (error) { qrStatus.textContent = error.message || "Immagine non leggibile."; }
  }

  async function findBolognaTrees(treeNumber) {
    const id = normalizeTreeIdentifier(treeNumber).replace(/'/g, "''");
    const where = encodeURIComponent(`num_pt='${id}' OR cod_alb='${id}'`);
    const url = `https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets/alberi-manutenzioni/records?where=${where}&limit=100`;
    const response = await fetch(url, { headers: { Accept: "application/json" } });
    if (!response.ok) throw new Error(`Servizio comunale non disponibile (${response.status}).`);
    const payload = await response.json();
    if (!payload.results?.length) return [];
    const records = [...payload.results];
    const maximumMatches = Math.min(Number(payload.total_count) || records.length, 1000);
    for (let offset = 100; offset < maximumMatches; offset += 100) {
      const nextResponse = await fetch(`${url}&offset=${offset}`, { headers: { Accept: "application/json" } });
      if (!nextResponse.ok) throw new Error(`Servizio comunale non disponibile (${nextResponse.status}).`);
      const nextPayload = await nextResponse.json();
      records.push(...(nextPayload.results || []));
    }
    return records.filter((item) => normalizeTreeIdentifier(item.num_pt) === id || normalizeTreeIdentifier(item.cod_alb) === id);
  }

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
    if (style === "hybrid") hybridLabels = L.tileLayer(TILE_LAYERS.labels.url, TILE_LAYERS.labels.options).addTo(map);
  }

  function initializeMap() {
    if (map || !window.L) return;
    map = L.map(mapNode, { zoomControl: true, zoomAnimation: false, fadeAnimation: false, markerZoomAnimation: false }).setView([44.4949, 11.3426], 13);
    treesLayer = L.layerGroup().addTo(map);
    applyMapStyle(mapStyle?.value || "classic");
    map.on("moveend zoomend", scheduleVisibleTrees);
  }

  function geolocationErrorMessage(error) {
    if (error?.code === 1) return "Permesso posizione negato. Abilita la posizione nelle impostazioni del dispositivo e riprova.";
    if (error?.code === 2) return "Posizione non disponibile. Controlla che il GPS sia attivo e riprova.";
    if (error?.code === 3) return "Ricerca della posizione scaduta. Riprova in un punto con migliore segnale GPS.";
    return error?.message || "Impossibile trovare la tua posizione.";
  }

  function showUserLocation(position) {
    initializeMap();
    if (!map) throw new Error("Mappa non disponibile.");
    const lat = Number(position?.coords?.latitude);
    const lng = Number(position?.coords?.longitude);
    const accuracy = Number(position?.coords?.accuracy);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) throw new Error("Coordinate della posizione non valide.");
    const point = L.latLng(lat, lng);
    if (!userLocationMarker) {
      userLocationMarker = L.circleMarker(point, {
        radius: 9,
        color: "#ffffff",
        weight: 3,
        fillColor: "#1268e8",
        fillOpacity: 1,
        pane: "markerPane"
      }).addTo(map).bindPopup("<strong>La mia posizione</strong>");
    } else {
      userLocationMarker.setLatLng(point).addTo(map);
    }
    if (Number.isFinite(accuracy) && accuracy > 0) {
      if (!userAccuracyCircle) {
        userAccuracyCircle = L.circle(point, {
          radius: accuracy,
          color: "#1268e8",
          weight: 1,
          opacity: 0.65,
          fillColor: "#4a9bff",
          fillOpacity: 0.14,
          interactive: false
        }).addTo(map);
      } else {
        userAccuracyCircle.setLatLng(point).setRadius(accuracy).addTo(map);
      }
      userLocationMarker.setPopupContent(`<strong>La mia posizione</strong><br>Precisione circa ${Math.round(accuracy)} m`);
    } else {
      userAccuracyCircle?.remove();
      userAccuracyCircle = null;
      userLocationMarker.setPopupContent("<strong>La mia posizione</strong>");
    }
    userLocationMarker.bringToFront().openPopup();
    map.setView(point, Math.max(map.getZoom(), 18), { animate: false });
    resizeMap();
    mapStatus.textContent = Number.isFinite(accuracy) && accuracy > 0
      ? `Posizione trovata (precisione circa ${Math.round(accuracy)} m).`
      : "Posizione trovata e mostrata sulla mappa.";
  }

  function centerOnUserLocation() {
    if (!navigator.geolocation) {
      mapStatus.textContent = "La posizione GPS non è supportata da questo dispositivo.";
      return;
    }
    mapLocationButton.disabled = true;
    mapLocationButton.setAttribute("aria-busy", "true");
    mapLocationButton.textContent = "⌖ RICERCA…";
    mapStatus.textContent = "Ricerca della tua posizione…";
    navigator.geolocation.getCurrentPosition(
      (position) => {
        try {
          showUserLocation(position);
        } catch (error) {
          mapStatus.textContent = geolocationErrorMessage(error);
        } finally {
          mapLocationButton.disabled = false;
          mapLocationButton.removeAttribute("aria-busy");
          mapLocationButton.textContent = "⌖ LA MIA POSIZIONE";
        }
      },
      (error) => {
        mapStatus.textContent = geolocationErrorMessage(error);
        mapLocationButton.disabled = false;
        mapLocationButton.removeAttribute("aria-busy");
        mapLocationButton.textContent = "⌖ LA MIA POSIZIONE";
      },
      { enableHighAccuracy: true, timeout: 12000, maximumAge: 30000 }
    );
  }

  function scheduleVisibleTrees() {
    clearTimeout(viewportTimer);
    viewportTimer = setTimeout(loadVisibleTrees, 350);
  }

  function distanceMeters(a, b) {
    const toRad = (value) => value * Math.PI / 180;
    const dLat = toRad(b.lat - a.lat);
    const dLng = toRad(b.lng - a.lng);
    const x = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(a.lat)) * Math.cos(toRad(b.lat)) * Math.sin(dLng / 2) ** 2;
    return Math.round(6371000 * 2 * Math.atan2(Math.sqrt(x), Math.sqrt(1 - x)));
  }

  function numberedTreeIcon(tree) {
    const treeNumber = esc(tree.num_pt || tree.cod_alb || "?");
    return L.divIcon({
      className: "tree-number-marker-wrap",
      html: `<span class="tree-number-marker">${treeNumber}</span>`,
      iconSize: [46, 28],
      iconAnchor: [23, 14],
      popupAnchor: [0, -15]
    });
  }

  function addVisibleTree(tree, targetLayer = treesLayer) {
    const point = tree.geo_point_2d;
    if (!point || !Number.isFinite(point.lat) || !Number.isFinite(point.lon)) return;
    const treeMarker = L.marker([point.lat, point.lon], { icon: numberedTreeIcon(tree), title: `Albero ${tree.num_pt || tree.cod_alb || ""}` });
    treeMarker.bindPopup(`<strong>${esc(tree.classe || "Specie non disponibile")}</strong><br>Numero: ${esc(tree.num_pt || "—")}<br>Codice: ${esc(tree.cod_alb || "—")}<br><button type="button" class="tree-popup-open" data-tree-number="${esc(tree.num_pt || tree.cod_alb)}">APRI SCHEDA</button>`);
    treeMarker.on("popupopen", (event) => {
      event.popup.getElement()?.querySelector(".tree-popup-open")?.addEventListener("click", () => showTree(tree), { once: true });
    });
    treeMarker.addTo(targetLayer);
  }

  async function loadVisibleTrees() {
    if (!map || page.classList.contains("hidden")) return;
    const zoom = map.getZoom();
    if (zoom < 16) {
      mapStatus.textContent = "Aumenta lo zoom almeno al livello 16. Gli alberi già caricati restano visibili.";
      return;
    }
    const center = map.getCenter();
    const viewportKey = `${zoom}:${center.lat.toFixed(4)}:${center.lng.toFixed(4)}`;
    if (viewportKey === lastViewportKey) return;
    const requestId = ++viewportRequest;
    viewportAbort?.abort();
    const controller = new AbortController();
    viewportAbort = controller;
    const radius = Math.min(1600, Math.max(80, distanceMeters(center, map.getBounds().getNorthEast()) + 40));
    mapStatus.textContent = "Aggiornamento della zona… Gli alberi attuali rimangono visibili.";
    try {
      const where = encodeURIComponent(`within_distance(geo_point_2d, geom'POINT(${center.lng} ${center.lat})', ${radius}m)`);
      const firstUrl = `https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets/alberi-manutenzioni/records?where=${where}&limit=100`;
      const firstResponse = await fetch(firstUrl, { headers: { Accept: "application/json" }, signal: controller.signal });
      if (!firstResponse.ok) throw new Error(`Servizio comunale non disponibile (${firstResponse.status}).`);
      const first = await firstResponse.json();
      if (requestId !== viewportRequest) return;
      if (first.total_count > 500) {
        mapStatus.textContent = `${first.total_count} alberi in questa zona: aumenta ancora lo zoom per visualizzare tutti i numeri.`;
        return;
      }
      const records = [...(first.results || [])];
      for (let offset = 100; offset < first.total_count; offset += 100) {
        const response = await fetch(`${firstUrl}&offset=${offset}`, { headers: { Accept: "application/json" }, signal: controller.signal });
        if (!response.ok) throw new Error(`Servizio comunale non disponibile (${response.status}).`);
        const payload = await response.json();
        if (requestId !== viewportRequest) return;
        records.push(...(payload.results || []));
      }
      const bounds = map.getBounds().pad(0.05);
      const nextLayer = L.layerGroup();
      records.filter((tree) => tree.geo_point_2d && bounds.contains([tree.geo_point_2d.lat, tree.geo_point_2d.lon])).forEach((tree) => addVisibleTree(tree, nextLayer));
      if (requestId !== viewportRequest) return;
      treesLayer?.remove();
      treesLayer = nextLayer.addTo(map);
      lastViewportKey = viewportKey;
      mapStatus.textContent = `${treesLayer.getLayers().length} alberi visualizzati. Tocca un numero per aprire la scheda.`;
    } catch (error) {
      if (error?.name === "AbortError") return;
      if (requestId === viewportRequest) mapStatus.textContent = `${error.message || "Impossibile aggiornare la zona."} Gli alberi precedenti restano disponibili.`;
    }
  }

  function showTree(tree) {
    const point = tree.geo_point_2d;
    if (!point || !Number.isFinite(point.lat) || !Number.isFinite(point.lon)) throw new Error("Albero trovato, ma senza coordinate utilizzabili.");
    const details = buildTreeDetails(tree);
    result.innerHTML = `<div class="tree-result-title"><div><small>Comune di Bologna · censimento ufficiale</small><h2>${esc(tree.classe || tree.nome_comune || "Specie non disponibile")}</h2></div><strong>#${esc(tree.num_pt || tree.cod_alb)}</strong></div><div class="tree-result-grid">${details}</div><p class="tree-data-source-note">Sono mostrati tutti i campi valorizzati restituiti in questo momento dal dataset ufficiale del Comune di Bologna. I campi vuoti non vengono visualizzati.</p><a class="btn btn-primary tree-navigate" href="https://www.google.com/maps/dir/?api=1&destination=${point.lat},${point.lon}" target="_blank" rel="noopener">NAVIGA VERSO L’ALBERO</a>`;
    result.classList.remove("hidden");
    initializeMap();
    if (marker) marker.remove();
    marker = L.marker([point.lat, point.lon]).addTo(map).bindPopup(`<strong>${esc(tree.classe || "Albero")}</strong><br>Numero ${esc(tree.num_pt || tree.cod_alb)}`).openPopup();
    map.setView([point.lat, point.lon], 19);
    setTimeout(() => { map.invalidateSize(); loadVisibleTrees(); }, 50);
  }

  function showTreeMatches(trees, identifier) {
    initializeMap();
    if (marker) marker.remove();
    marker = null;
    const nextLayer = L.layerGroup();
    const points = [];
    trees.forEach((tree) => {
      const point = tree.geo_point_2d;
      if (!point || !Number.isFinite(point.lat) || !Number.isFinite(point.lon)) return;
      points.push([point.lat, point.lon]);
      addVisibleTree(tree, nextLayer);
    });
    if (!points.length) throw new Error("Alberi trovati, ma senza coordinate utilizzabili.");
    treesLayer?.remove();
    treesLayer = nextLayer.addTo(map);
    lastViewportKey = "";
    result.innerHTML = `<div class="tree-result-title"><div><small>Comune di Bologna</small><h2>Codice albero ${esc(identifier)}</h2></div><strong>${trees.length}</strong></div><p>Questo codice è associato a più alberi. Sono tutti indicati sulla mappa: tocca il numero della pianta per aprire la scheda corretta.</p>`;
    result.classList.remove("hidden");
    map.fitBounds(L.latLngBounds(points).pad(0.08), { animate: false, maxZoom: 18 });
    mapStatus.textContent = `${points.length} alberi con codice ${identifier}. Tocca un numero per aprire la scheda.`;
    resizeMap();
  }

  form?.addEventListener("submit", async (event) => {
    event.preventDefault();
    setStatus("Ricerca nel censimento ufficiale…");
    result.classList.add("hidden");
    try {
      if (municipality.value !== "Bologna") throw new Error("Il censimento di questo Comune non è ancora collegato.");
      const identifier = normalizeTreeIdentifier(number.value);
      number.value = identifier;
      const trees = await findBolognaTrees(identifier);
      if (!trees.length) throw new Error("Albero non trovato. Controlla il numero punto o il codice albero riportato sul cartellino.");
      const pointMatch = trees.find((tree) => normalizeTreeIdentifier(tree.num_pt) === identifier);
      if (pointMatch) {
        showTree(pointMatch);
        setStatus("Albero trovato tramite numero punto nel censimento del Comune di Bologna.", "success");
      } else if (trees.length === 1) {
        showTree(trees[0]);
        setStatus("Albero trovato tramite codice albero nel censimento del Comune di Bologna.", "success");
      } else {
        showTreeMatches(trees, identifier);
        setStatus(`${trees.length} alberi trovati con il codice ${identifier}. Seleziona il numero corretto sulla mappa.`, "success");
      }
    } catch (error) { setStatus(error.message || "Ricerca non riuscita.", "error"); }
  });
  $("open-tree-search-btn")?.addEventListener("click", openPage);
  $("tree-search-back-btn")?.addEventListener("click", closePage);
  $("tree-qr-open-btn")?.addEventListener("click", startScanner);
  mapLocationButton?.addEventListener("click", centerOnUserLocation);
  mapFullscreenButton?.addEventListener("click", () => setMapFullscreen(!mapFullscreen));
  mapStyle?.addEventListener("change", () => applyMapStyle(mapStyle.value));
  $("tree-qr-close-btn")?.addEventListener("click", stopScanner);
  dialog?.addEventListener("close", stopScanner);
  $("tree-qr-file")?.addEventListener("change", (event) => event.target.files?.[0] && scanFile(event.target.files[0]));
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && mapFullscreen && !dialog?.open) setMapFullscreen(false);
  });
})();

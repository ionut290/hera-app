(() => {
  "use strict";
  const $ = (id) => document.getElementById(id);
  const page = $("tree-search-page");
  const form = $("tree-search-form");
  const municipality = $("tree-municipality");
  const number = $("tree-number");
  const status = $("tree-search-status");
  const result = $("tree-result");
  const mapNode = $("tree-map");
  const mapStatus = $("tree-map-status");
  const mapStyle = $("tree-map-style");
  const dialog = $("tree-qr-dialog");
  const video = $("tree-qr-video");
  const qrStatus = $("tree-qr-status");
  let map = null;
  let marker = null;
  let treesLayer = null;
  let baseLayer = null;
  let hybridLabels = null;
  let viewportTimer = 0;
  let viewportRequest = 0;
  let stream = null;
  let scanFrame = 0;

  const esc = (value) => String(value ?? "—").replace(/[&<>'"]/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[char]));
  const setStatus = (message, type = "") => { status.textContent = message; status.className = `tree-search-status ${type}`.trim(); };
  const hideHome = (hidden) => $("home-page")?.classList.toggle("hidden", hidden);

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

  async function findBolognaTree(treeNumber) {
    const id = String(treeNumber).trim().replace(/'/g, "''");
    const where = encodeURIComponent(`num_pt='${id}' OR cod_alb='${id}'`);
    const url = `https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets/alberi-manutenzioni/records?where=${where}&limit=10`;
    const response = await fetch(url, { headers: { Accept: "application/json" } });
    if (!response.ok) throw new Error(`Servizio comunale non disponibile (${response.status}).`);
    const payload = await response.json();
    if (!payload.results?.length) return null;
    return payload.results.find((item) => String(item.num_pt) === String(treeNumber) || String(item.cod_alb).toLowerCase() === String(treeNumber).toLowerCase()) || payload.results[0];
  }

  const TILE_LAYERS = {
    classic: { url: "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", options: { maxZoom: 20, attribution: "&copy; OpenStreetMap" } },
    satellite: { url: "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, attribution: "Tiles &copy; Esri" } },
    labels: { url: "https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, attribution: "Labels &copy; Esri", pane: "overlayPane" } }
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
    map = L.map(mapNode, { zoomControl: true }).setView([44.4949, 11.3426], 13);
    treesLayer = L.layerGroup().addTo(map);
    applyMapStyle(mapStyle?.value || "classic");
    map.on("moveend zoomend", scheduleVisibleTrees);
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

  function addVisibleTree(tree) {
    const point = tree.geo_point_2d;
    if (!point || !Number.isFinite(point.lat) || !Number.isFinite(point.lon)) return;
    const treeMarker = L.marker([point.lat, point.lon], { icon: numberedTreeIcon(tree), title: `Albero ${tree.num_pt || tree.cod_alb || ""}` });
    treeMarker.bindPopup(`<strong>${esc(tree.classe || "Specie non disponibile")}</strong><br>Numero: ${esc(tree.num_pt || "—")}<br>Codice: ${esc(tree.cod_alb || "—")}<br><button type="button" class="tree-popup-open" data-tree-number="${esc(tree.num_pt || tree.cod_alb)}">APRI SCHEDA</button>`);
    treeMarker.on("popupopen", (event) => {
      event.popup.getElement()?.querySelector(".tree-popup-open")?.addEventListener("click", () => showTree(tree), { once: true });
    });
    treeMarker.addTo(treesLayer);
  }

  async function loadVisibleTrees() {
    if (!map || page.classList.contains("hidden")) return;
    const zoom = map.getZoom();
    treesLayer?.clearLayers();
    if (zoom < 16) {
      mapStatus.textContent = "Aumenta lo zoom almeno al livello 16 per visualizzare i numeri degli alberi.";
      return;
    }
    const requestId = ++viewportRequest;
    const center = map.getCenter();
    const radius = Math.min(1600, Math.max(80, distanceMeters(center, map.getBounds().getNorthEast()) + 40));
    mapStatus.textContent = "Caricamento degli alberi nella zona visibile…";
    try {
      const where = encodeURIComponent(`within_distance(geo_point_2d, geom'POINT(${center.lng} ${center.lat})', ${radius}m)`);
      const firstUrl = `https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets/alberi-manutenzioni/records?where=${where}&limit=100`;
      const firstResponse = await fetch(firstUrl, { headers: { Accept: "application/json" } });
      if (!firstResponse.ok) throw new Error(`Servizio comunale non disponibile (${firstResponse.status}).`);
      const first = await firstResponse.json();
      if (requestId !== viewportRequest) return;
      if (first.total_count > 500) {
        mapStatus.textContent = `${first.total_count} alberi in questa zona: aumenta ancora lo zoom per visualizzare tutti i numeri.`;
        return;
      }
      const records = [...(first.results || [])];
      for (let offset = 100; offset < first.total_count; offset += 100) {
        const response = await fetch(`${firstUrl}&offset=${offset}`, { headers: { Accept: "application/json" } });
        if (!response.ok) throw new Error(`Servizio comunale non disponibile (${response.status}).`);
        const payload = await response.json();
        if (requestId !== viewportRequest) return;
        records.push(...(payload.results || []));
      }
      const bounds = map.getBounds().pad(0.05);
      records.filter((tree) => tree.geo_point_2d && bounds.contains([tree.geo_point_2d.lat, tree.geo_point_2d.lon])).forEach(addVisibleTree);
      mapStatus.textContent = `${treesLayer.getLayers().length} alberi visualizzati. Tocca un numero per aprire la scheda.`;
    } catch (error) {
      if (requestId === viewportRequest) mapStatus.textContent = error.message || "Impossibile caricare gli alberi della zona.";
    }
  }

  function showTree(tree) {
    const point = tree.geo_point_2d;
    if (!point || !Number.isFinite(point.lat) || !Number.isFinite(point.lon)) throw new Error("Albero trovato, ma senza coordinate utilizzabili.");
    result.innerHTML = `<div class="tree-result-title"><div><small>Comune di Bologna</small><h2>${esc(tree.classe || "Specie non disponibile")}</h2></div><strong>#${esc(tree.num_pt || tree.cod_alb)}</strong></div><div class="tree-result-grid"><div><span>Numero punto</span><strong>${esc(tree.num_pt)}</strong></div><div><span>Codice albero</span><strong>${esc(tree.cod_alb)}</strong></div><div><span>Altezza</span><strong>${esc(tree.cl_h)}</strong></div><div><span>Circonferenza</span><strong>${esc(tree.classe_circonferenza_diametro)}</strong></div><div><span>Quartiere</span><strong>${esc(tree.quartiere)}</strong></div><div><span>Albero di pregio</span><strong>${tree.pregio === "S" ? "Sì" : "No"}</strong></div></div><a class="btn btn-primary tree-navigate" href="https://www.google.com/maps/dir/?api=1&destination=${point.lat},${point.lon}" target="_blank" rel="noopener">NAVIGA VERSO L’ALBERO</a>`;
    result.classList.remove("hidden");
    initializeMap();
    if (marker) marker.remove();
    marker = L.marker([point.lat, point.lon]).addTo(map).bindPopup(`<strong>${esc(tree.classe || "Albero")}</strong><br>Numero ${esc(tree.num_pt || tree.cod_alb)}`).openPopup();
    map.setView([point.lat, point.lon], 19);
    setTimeout(() => { map.invalidateSize(); loadVisibleTrees(); }, 50);
  }

  form?.addEventListener("submit", async (event) => {
    event.preventDefault();
    setStatus("Ricerca nel censimento ufficiale…");
    result.classList.add("hidden");
    try {
      if (municipality.value !== "Bologna") throw new Error("Il censimento di questo Comune non è ancora collegato.");
      const tree = await findBolognaTree(number.value);
      if (!tree) throw new Error("Albero non trovato. Controlla Comune e numero riportati sul cartellino.");
      showTree(tree); setStatus("Albero trovato nel censimento del Comune di Bologna.", "success");
    } catch (error) { setStatus(error.message || "Ricerca non riuscita.", "error"); }
  });
  $("open-tree-search-btn")?.addEventListener("click", openPage);
  $("tree-search-back-btn")?.addEventListener("click", closePage);
  $("tree-qr-open-btn")?.addEventListener("click", startScanner);
  mapStyle?.addEventListener("change", () => applyMapStyle(mapStyle.value));
  $("tree-qr-close-btn")?.addEventListener("click", stopScanner);
  dialog?.addEventListener("close", stopScanner);
  $("tree-qr-file")?.addEventListener("change", (event) => event.target.files?.[0] && scanFile(event.target.files[0]));
})();

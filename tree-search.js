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
  let mapNode = $("tree-map");
  const mapStatus = $("tree-map-status");
  const mapStyle = $("tree-map-style");
  const mapPlantingFilter = $("tree-map-planting-filter");
  const mapLocationButton = $("tree-map-location-btn");
  const mapFullscreenButton = $("tree-map-fullscreen-btn");
  const dialog = $("tree-qr-dialog");
  const video = $("tree-qr-video");
  const qrStatus = $("tree-qr-status");
  const VERDE_BOLOGNA_RETURN_KEY = "varga-verde-bologna:return-from-tree";
  const MOBILE_QUERY = "(max-width: 760px)";
  const MOBILE_VISIBLE_TREE_LIMIT = 80;
  const DESKTOP_VISIBLE_TREE_LIMIT = 300;
  const MOBILE_MAP_START_DELAY_MS = 250;
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
  let pageActivation = 0;
  let mapStartTimer = 0;

  const esc = (value) => String(value ?? "—").replace(/[&<>'"]/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[char]));
  const setStatus = (message, type = "") => { status.textContent = message; status.className = `tree-search-status ${type}`.trim(); };
  const hideHome = (hidden) => $("home-page")?.classList.toggle("hidden", hidden);
  const normalizeTreeIdentifier = (value) => String(value ?? "").trim().replace(/\s+/g, "").toUpperCase();
  const mobileView = () => window.matchMedia?.(MOBILE_QUERY)?.matches === true;

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
    irriga: "Irrigazione",
    distanza_fabbricati: "Distanza dai fabbricati",
    d_edif: "Distanza dai fabbricati",
    data_impianto: "Data impianto",
    data_impnt: "Data impianto",
    data_inventario: "Data inventario",
    data_inv: "Data inserimento inventario",
    data_aggiornamento: "Data aggiornamento",
    data_agg: "Data ultimo aggiornamento",
    anni_impnt: "Anni dall’impianto",
    stato: "Stato",
    note: "Note",
    geo_point_2d: "Coordinate"
  };

  const TREE_FIELD_PRIORITY = [
    "num_pt", "cod_alb", "classe", "nome_scientifico", "nome_comune",
    "cl_h", "classe_circonferenza_diametro", "circonferenza", "diametro", "altezza",
    "quartiere", "via", "indirizzo", "localizzazione", "dimora", "pregio", "irrigazione",
    "irriga", "distanza_fabbricati", "d_edif", "data_impianto", "data_impnt", "data_inventario", "data_inv",
    "data_aggiornamento", "data_agg", "anni_impnt", "stato", "note", "geo_point_2d"
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
    if (["data_impianto", "data_impnt", "data_inventario", "data_inv", "data_aggiornamento", "data_agg"].includes(key)) {
      const normalized = String(value || "").slice(0, 10);
      const match = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (match) return `${match[3]}/${match[2]}/${match[1]}`;
    }
    if (Array.isArray(value)) return value.map((item) => typeof item === "object" ? JSON.stringify(item) : String(item)).join(", ");
    if (typeof value === "object") {
      try { return JSON.stringify(value); } catch (_) { return String(value); }
    }
    return String(value);
  }

  function buildTreeDetailEntries(tree) {
    const keys = Object.keys(tree || {}).filter((key) => hasTreeValue(tree[key]));
    const ordered = [
      ...TREE_FIELD_PRIORITY.filter((key) => keys.includes(key)),
      ...keys.filter((key) => !TREE_FIELD_PRIORITY.includes(key)).sort((a, b) => treeFieldLabel(a).localeCompare(treeFieldLabel(b), "it"))
    ];
    return ordered.map((key) => ({
      key,
      label: treeFieldLabel(key),
      value: formatTreeValue(key, tree[key])
    }));
  }

  function buildTreeDetails(entries) {
    return entries.map((entry, index) => `<div class="${index >= 6 ? "tree-detail-extra" : ""}"${index >= 6 ? " hidden" : ""}><span>${esc(entry.label)}</span><strong>${esc(entry.value)}</strong></div>`).join("");
  }

  function buildTreeWhazzupMessage(entries, navigationUrl) {
    const details = entries.slice(0, 6).map((entry) => `• *${entry.label}:* ${entry.value}`);
    return [
      "🌳 *SCHEDA ALBERO*",
      "",
      ...details,
      "",
      "📍 *NAVIGA VERSO L’ALBERO*",
      navigationUrl
    ].join("\n");
  }

  async function openTreeShareInWhazzup(message) {
    const appUrl = `whatsapp://send?text=${encodeURIComponent(String(message || ""))}`;
    const nativeAndroid = Boolean(
      window.Capacitor?.isNativePlatform?.()
      && window.Capacitor?.getPlatform?.() === "android"
    );
    if (nativeAndroid) {
      const plugin = window.Capacitor?.Plugins?.HeraWhatsApp
        || window.Capacitor?.registerPlugin?.("HeraWhatsApp")
        || null;
      if (!plugin?.open) {
        window.alert("WhatsApp non è disponibile su questo dispositivo.");
        return false;
      }
      try {
        const response = await plugin.open({ url: appUrl });
        return response?.opened !== false;
      } catch (error) {
        window.alert(error?.message || "WhatsApp non è installato o non può essere aperto su questo dispositivo.");
        return false;
      }
    }
    window.location.assign(appUrl);
    window.setTimeout(() => {
      if (document.visibilityState === "visible") {
        window.alert("WhatsApp non è installato o non può essere aperto su questo dispositivo.");
      }
    }, 1800);
    return true;
  }

  function waitForTreeStreetViewCards(timeoutMs = 3000) {
    if (typeof window.HeraStreetViewCards?.openForCoordinates === "function") {
      return Promise.resolve(window.HeraStreetViewCards);
    }
    return new Promise((resolve) => {
      const startedAt = Date.now();
      const timer = window.setInterval(() => {
        const api = window.HeraStreetViewCards;
        if (typeof api?.openForCoordinates === "function" || Date.now() - startedAt >= timeoutMs) {
          window.clearInterval(timer);
          resolve(typeof api?.openForCoordinates === "function" ? api : null);
        }
      }, 100);
    });
  }

  async function openTreeStreetView(tree, point, button) {
    const api = await waitForTreeStreetViewCards();
    if (!api) {
      window.alert("Vista panoramica 360° non disponibile in questo momento. Riprova tra qualche secondo.");
      return false;
    }
    const treeNumber = tree.num_pt || tree.cod_alb || "";
    return api.openForCoordinates(
      { lat: Number(point.lat), lng: Number(point.lon) },
      button,
      {
        targetLabel: "Albero",
        modalTitle: `🌐 Vista 360° albero${treeNumber ? ` #${treeNumber}` : ""}`
      }
    );
  }

  function safeExternalUrl(value) {
    try {
      const url = new URL(String(value || ""));
      return ["http:", "https:"].includes(url.protocol) ? url.href : "";
    } catch (_) { return ""; }
  }

  async function treeAssistantApi(action, payload = {}) {
    if (navigator.onLine === false) throw new Error("Questa funzione richiede una connessione Internet.");
    const user = window.firebase?.auth?.().currentUser;
    if (!user) throw new Error("Accedi all’app per consultare la manutenzione.");
    const token = await user.getIdToken();
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), 55000);
    try {
      const response = await fetch("/api/green-assistant", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
        body: JSON.stringify({ action, ...payload }),
        signal: controller.signal
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data.ok) throw new Error(data.error || "Servizio di manutenzione temporaneamente non disponibile.");
      return data.result ?? data;
    } catch (error) {
      if (error?.name === "AbortError") throw new Error("Il servizio sta impiegando troppo tempo. Riprova.");
      throw error;
    } finally { window.clearTimeout(timeout); }
  }

  function treeMaintenancePayload(tree) {
    return {
      scientificName: tree.nome_scientifico || tree.classe || "",
      commonName: tree.nome_comune || tree.classe || "",
      heightClass: tree.cl_h || tree.altezza || "",
      diameter: tree.diametro || tree.circonferenza || tree.classe_circonferenza_diametro || "",
      plantingYear: tree.data_impianto || tree.data_impnt || tree.anni_impnt || "",
      location: [tree.via, tree.indirizzo, tree.quartiere, tree.localizzazione].filter(Boolean).join(" · "),
      irrigation: tree.irrigazione || tree.irriga || "",
      censusNotes: tree.note || tree.stato || ""
    };
  }

  function treeMaintenanceCacheKey(tree) {
    const payload = treeMaintenancePayload(tree);
    return `heraTreeMaintenanceV1:${String(payload.scientificName || payload.commonName).trim().toLowerCase()}`;
  }

  function readTreeMaintenanceCache(tree) {
    try {
      const cached = JSON.parse(localStorage.getItem(treeMaintenanceCacheKey(tree)) || "null");
      return cached && Date.now() - Number(cached.savedAt || 0) < 30 * 86400000 ? cached.value : null;
    } catch (_) { return null; }
  }

  function saveTreeMaintenanceCache(tree, value) {
    try { localStorage.setItem(treeMaintenanceCacheKey(tree), JSON.stringify({ savedAt: Date.now(), value })); } catch (_) {}
  }

  function maintenanceList(title, items) {
    const values = Array.isArray(items) ? items : [];
    return values.length ? `<section class="tree-maintenance-section"><h3>${esc(title)}</h3><ul>${values.map((item) => `<li>${esc(item)}</li>`).join("")}</ul></section>` : "";
  }

  function renderTreeMaintenance(tree, data, cached = false) {
    const panel = result.querySelector(".tree-maintenance-panel");
    if (!panel) return;
    const maintenance = Array.isArray(data?.maintenance) ? data.maintenance : [];
    const diseases = Array.isArray(data?.commonDiseases) ? data.commonDiseases : [];
    const photos = Array.isArray(data?.photos) ? data.photos : [];
    const sources = Array.isArray(data?.sources) ? data.sources : [];
    panel.innerHTML = `<div class="tree-maintenance-head"><div><small>${cached ? "Scheda salvata sul dispositivo" : "Gemini · fonti Brave · fotografie iNaturalist"}</small><h2>🪚 Manutenzione di ${esc(data?.species || tree.classe || "questo albero")}</h2></div><button class="btn tree-maintenance-close" type="button">CHIUDI</button></div>
      <p class="tree-maintenance-warning">⚠️ ${esc(data?.notice || data?.warning || "Informazioni orientative da verificare sul posto.")}</p>
      ${data?.summary ? `<p>${esc(data.summary)}</p>` : ""}
      ${maintenance.length ? `<section class="tree-maintenance-section"><h3>Calendario degli interventi</h3><div class="tree-maintenance-table">${maintenance.map((item) => `<article><strong>${esc(item.intervention)}</strong><span><b>Periodo:</b> ${esc(item.period || "Da valutare")}</span><span><b>Frequenza:</b> ${esc(item.frequency || "Secondo necessità")}</span><p>${esc(item.notes || "")}</p></article>`).join("")}</div></section>` : ""}
      ${maintenanceList("Irrigazione", data?.watering)}
      ${maintenanceList("Potatura", data?.pruning)}
      ${maintenanceList("Controlli periodici", data?.inspections)}
      ${diseases.length ? `<section class="tree-maintenance-section"><h3>Malattie e problemi frequenti della specie</h3><div class="tree-disease-grid">${diseases.map((item) => `<article><strong>${esc(item.name)}</strong><p><b>Segnali:</b> ${esc(item.symptoms)}</p><p><b>Azione prudente:</b> ${esc(item.action)}</p></article>`).join("")}</div></section>` : ""}
      ${maintenanceList("Sicurezza", data?.safety)}
      ${photos.length ? `<section class="tree-maintenance-section"><h3>Fotografie di riferimento</h3><div class="tree-reference-photos">${photos.map((item) => { const image = safeExternalUrl(item.image); const url = safeExternalUrl(item.url); return image && url ? `<a href="${esc(url)}" target="_blank" rel="noopener"><img src="${esc(image)}" alt="${esc(item.name || data.species)}" loading="lazy"><small>${esc(item.attribution || "iNaturalist")}${item.license ? ` · ${esc(item.license)}` : ""}</small></a>` : ""; }).join("")}</div></section>` : ""}
      ${sources.length ? `<section class="tree-maintenance-section"><h3>Fonti da consultare</h3><div class="tree-maintenance-sources">${sources.map((item) => { const url = safeExternalUrl(item.url); return url ? `<a href="${esc(url)}" target="_blank" rel="noopener"><strong>${esc(item.title || item.domain)}</strong><small>${esc(item.domain)}</small><span>${esc(item.description || "")}</span></a>` : ""; }).join("")}</div></section>` : ""}
      <section class="tree-maintenance-section tree-photo-diagnosis"><h3>📷 Controlla una possibile malattia</h3><p>Fotografa la parte danneggiata. Pl@ntNet confronterà i sintomi visibili; il risultato non è una diagnosi definitiva.</p><label class="btn tree-disease-photo-label">SCEGLI FOTO<input class="tree-disease-photo" type="file" accept="image/jpeg,image/png" capture="environment"></label><div class="tree-disease-photo-result" role="status" aria-live="polite"></div></section>
      <p class="tree-maintenance-warning">${esc(data?.warning || "Per rischi strutturali, sintomi gravi o interventi importanti richiedere un arboricoltore qualificato.")}</p>`;
    panel.classList.remove("hidden");
    panel.querySelector(".tree-maintenance-close")?.addEventListener("click", () => panel.classList.add("hidden"));
    panel.querySelector(".tree-disease-photo")?.addEventListener("change", (event) => analyzeTreeDiseasePhoto(event.currentTarget));
    panel.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function compressTreePhoto(file) {
    if (!file || !/^image\/(jpeg|png)$/i.test(file.type)) throw new Error("Seleziona una fotografia JPG o PNG.");
    const objectUrl = URL.createObjectURL(file);
    try {
      const image = await new Promise((resolve, reject) => { const img = new Image(); img.onload = () => resolve(img); img.onerror = () => reject(new Error("Impossibile leggere la fotografia.")); img.src = objectUrl; });
      const scale = Math.min(1, 1600 / Math.max(image.naturalWidth, image.naturalHeight));
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round(image.naturalWidth * scale)); canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
      canvas.getContext("2d", { alpha: false }).drawImage(image, 0, 0, canvas.width, canvas.height);
      return canvas.toDataURL("image/jpeg", 0.82);
    } finally { URL.revokeObjectURL(objectUrl); }
  }

  async function analyzeTreeDiseasePhoto(input) {
    const target = result.querySelector(".tree-disease-photo-result");
    if (!target || !input.files?.[0]) return;
    target.textContent = "Analizzo la fotografia con Pl@ntNet…";
    input.disabled = true;
    try {
      const data = await treeAssistantApi("identifyDisease", { image: await compressTreePhoto(input.files[0]), organ: "auto" });
      const matches = Array.isArray(data?.results) ? data.results : [];
      target.innerHTML = matches.length ? `<div class="tree-disease-photo-matches">${matches.map((item) => `<article><strong>${esc(item.name || item.code || "Possibile problema")}</strong><span>Compatibilità visiva: ${Math.round(Number(item.score || 0) * 100)}%</span>${safeExternalUrl(item.image) ? `<img src="${esc(safeExternalUrl(item.image))}" alt="Immagine di confronto" loading="lazy">` : ""}</article>`).join("")}</div><p>Confronto fotografico orientativo: verifica sintomi, specie e condizioni sul posto.</p>` : "Nessuna corrispondenza affidabile trovata.";
    } catch (error) { target.textContent = error.message; }
    finally { input.disabled = false; input.value = ""; }
  }

  async function openTreeMaintenance(tree, button) {
    const cached = readTreeMaintenanceCache(tree);
    if (cached) return renderTreeMaintenance(tree, cached, true);
    const panel = result.querySelector(".tree-maintenance-panel");
    if (panel) { panel.classList.remove("hidden"); panel.innerHTML = "<p>Preparo manutenzione, controlli, malattie, fotografie e fonti…</p>"; }
    button.disabled = true;
    try {
      const data = await treeAssistantApi("treeMaintenance", treeMaintenancePayload(tree));
      saveTreeMaintenanceCache(tree, data);
      renderTreeMaintenance(tree, data);
    } catch (error) { if (panel) panel.innerHTML = `<p class="tree-maintenance-error">${esc(error.message)}</p>`; }
    finally { button.disabled = false; }
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
    const activation = ++pageActivation;
    window.clearTimeout(mapStartTimer);
    mapStartTimer = 0;
    document.getElementById("menu-close-btn")?.click();
    hideHome(true);
    page.classList.remove("hidden");
    page.setAttribute("aria-hidden", "false");
    page.scrollTo?.({ top: 0, behavior: "auto" });
    mapStatus.textContent = "Mappa pronta. Aumenta lo zoom per visualizzare i numeri degli alberi.";
    requestAnimationFrame(() => {
      mapStartTimer = window.setTimeout(() => {
        mapStartTimer = 0;
        if (activation !== pageActivation || page.classList.contains("hidden")) return;
        const currentMap = initializeMap();
        currentMap?.invalidateSize?.({ pan: false, animate: false });
        loadVisibleTrees();
      }, mobileView() ? MOBILE_MAP_START_DELAY_MS : 0);
    });
    if (!mobileView()) {
      window.setTimeout(() => {
        if (activation === pageActivation && !page.classList.contains("hidden")) number?.focus?.({ preventScroll: true });
      }, 0);
    }
  }

  function destroyTreeMap() {
    pageActivation += 1;
    window.clearTimeout(mapStartTimer);
    mapStartTimer = 0;
    window.clearTimeout(viewportTimer);
    viewportTimer = 0;
    viewportRequest += 1;
    viewportAbort?.abort?.();
    viewportAbort = null;
    try { map?.off?.(); map?.remove?.(); } catch (_) {}
    map = null;
    marker = null;
    treesLayer = null;
    baseLayer = null;
    hybridLabels = null;
    userLocationMarker = null;
    userAccuracyCircle = null;
    lastViewportKey = "";
    if (mapNode) { mapNode.innerHTML = ""; mapNode.removeAttribute("style"); try { delete mapNode._leaflet_id; } catch (_) {} mapNode.className = "tree-map"; }
  }

  function closePage() {
    stopScanner();
    setMapFullscreen(false);
    destroyTreeMap();
    page.classList.add("hidden");
    page.setAttribute("aria-hidden", "true");
    let returnToVerdeBologna = false;
    try { returnToVerdeBologna = sessionStorage.getItem(VERDE_BOLOGNA_RETURN_KEY) === "1"; sessionStorage.removeItem(VERDE_BOLOGNA_RETURN_KEY); } catch (_) {}
    if (returnToVerdeBologna && typeof window.HeraVerdeBologna?.open === "function") { window.HeraVerdeBologna.open(); return; }
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

  function renewTreeMapContainer() {
    if (!mapNode?.parentNode) return mapNode;
    const freshMapNode = mapNode.cloneNode(false);
    freshMapNode.className = "tree-map";
    freshMapNode.removeAttribute("style");
    freshMapNode.removeAttribute("tabindex");
    mapNode.replaceWith(freshMapNode);
    mapNode = freshMapNode;
    return mapNode;
  }

  function initializeMap() {
    if (map) return map;
    if (!window.L || !mapNode) {
      mapStatus.textContent = "Mappa non disponibile in questo momento. La ricerca per numero resta utilizzabile.";
      return null;
    }
    const createMap = () => {
      map = L.map(mapNode, { zoomControl: true, zoomAnimation: false, fadeAnimation: false, markerZoomAnimation: false }).setView([44.4949, 11.3426], 13);
      treesLayer = L.layerGroup().addTo(map);
      applyMapStyle(mapStyle?.value || "classic");
      map.on("moveend zoomend", scheduleVisibleTrees);
      return map;
    };
    try {
      return createMap();
    } catch (error) {
      try { map?.remove?.(); } catch (_) {}
      map = null;
      treesLayer = null;
      baseLayer = null;
      hybridLabels = null;
      renewTreeMapContainer();
      try {
        return createMap();
      } catch (retryError) {
        try { map?.remove?.(); } catch (_) {}
        map = null;
        treesLayer = null;
        mapStatus.textContent = `Impossibile inizializzare la mappa: ${retryError?.message || error?.message || "errore sconosciuto"}. La ricerca per numero resta utilizzabile.`;
        return null;
      }
    }
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

  function plantingFilterYears() {
    const years = Number(mapPlantingFilter?.value);
    return [1, 3, 5].includes(years) ? years : 0;
  }

  function formatApiDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function plantingDateClause(years) {
    if (!years) return "";
    const end = new Date();
    const start = new Date(end.getFullYear() - years, end.getMonth(), end.getDate());
    return `data_impnt >= date'${formatApiDate(start)}' AND data_impnt <= date'${formatApiDate(end)}'`;
  }

  function plantingFilterDescription(years) {
    if (years === 1) return "nuovi impianti nell’ultimo anno";
    if (years) return `nuovi impianti negli ultimi ${years} anni`;
    return "alberi";
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
    const filterYears = plantingFilterYears();
    const filterDescription = plantingFilterDescription(filterYears);
    if (zoom < 16) {
      mapStatus.textContent = `Aumenta lo zoom almeno al livello 16 per visualizzare ${filterDescription}. Gli alberi già caricati restano visibili.`;
      return;
    }
    const center = map.getCenter();
    const viewportKey = `${filterYears}:${zoom}:${center.lat.toFixed(4)}:${center.lng.toFixed(4)}`;
    if (viewportKey === lastViewportKey) return;
    const requestId = ++viewportRequest;
    viewportAbort?.abort();
    const controller = new AbortController();
    viewportAbort = controller;
    const radius = Math.min(1600, Math.max(80, distanceMeters(center, map.getBounds().getNorthEast()) + 40));
    mapStatus.textContent = "Aggiornamento della zona… Gli alberi attuali rimangono visibili.";
    try {
      const geographicClause = `within_distance(geo_point_2d, geom'POINT(${center.lng} ${center.lat})', ${radius}m)`;
      const where = encodeURIComponent([geographicClause, plantingDateClause(filterYears)].filter(Boolean).join(" AND "));
      const firstUrl = `https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets/alberi-manutenzioni/records?where=${where}&limit=100`;
      const firstResponse = await fetch(firstUrl, { headers: { Accept: "application/json" }, signal: controller.signal });
      if (!firstResponse.ok) throw new Error(`Servizio comunale non disponibile (${firstResponse.status}).`);
      const first = await firstResponse.json();
      if (requestId !== viewportRequest) return;
      const total = Number(first.total_count) || 0;
      const visibleLimit = mobileView() ? MOBILE_VISIBLE_TREE_LIMIT : DESKTOP_VISIBLE_TREE_LIMIT;
      const records = [...(first.results || [])].slice(0, visibleLimit);
      for (let offset = 100; offset < total && records.length < visibleLimit; offset += 100) {
        const response = await fetch(`${firstUrl}&offset=${offset}`, { headers: { Accept: "application/json" }, signal: controller.signal });
        if (!response.ok) throw new Error(`Servizio comunale non disponibile (${response.status}).`);
        const payload = await response.json();
        if (requestId !== viewportRequest) return;
        records.push(...(payload.results || []).slice(0, visibleLimit - records.length));
      }
      const bounds = map.getBounds().pad(0.05);
      const nextLayer = L.layerGroup();
      records.filter((tree) => tree.geo_point_2d && bounds.contains([tree.geo_point_2d.lat, tree.geo_point_2d.lon])).forEach((tree) => addVisibleTree(tree, nextLayer));
      if (requestId !== viewportRequest) return;
      treesLayer?.remove();
      treesLayer = nextLayer.addTo(map);
      lastViewportKey = viewportKey;
      const visibleCount = treesLayer.getLayers().length;
      const limitedNote = total > records.length ? ` su ${total} presenti nella zona; aumenta lo zoom per restringere` : "";
      mapStatus.textContent = `${visibleCount} ${filterDescription} visualizzati${limitedNote}. Tocca un numero per aprire la scheda.`;
    } catch (error) {
      if (error?.name === "AbortError") return;
      if (requestId === viewportRequest) mapStatus.textContent = `${error.message || "Impossibile aggiornare la zona."} Gli alberi precedenti restano disponibili.`;
    }
  }

  function showTree(tree) {
    const point = tree.geo_point_2d;
    if (!point || !Number.isFinite(point.lat) || !Number.isFinite(point.lon)) throw new Error("Albero trovato, ma senza coordinate utilizzabili.");
    const detailEntries = buildTreeDetailEntries(tree);
    const details = buildTreeDetails(detailEntries);
    const navigationUrl = `https://www.google.com/maps/dir/?api=1&destination=${point.lat},${point.lon}`;
    const expandButton = detailEntries.length > 6
      ? `<button class="btn tree-details-toggle" type="button" aria-expanded="false">MOSTRA TUTTI I DETTAGLI (${detailEntries.length})</button>`
      : "";
    const detailsNote = detailEntries.length > 6
      ? "Sono visibili i primi 6 dettagli. Premi il pulsante per consultare tutti i campi valorizzati del censimento ufficiale."
      : "Sono mostrati tutti i campi valorizzati disponibili nel censimento ufficiale.";
    result.innerHTML = `<div class="tree-result-title"><div><small>Comune di Bologna · censimento ufficiale</small><h2>${esc(tree.classe || tree.nome_comune || "Specie non disponibile")}</h2></div><strong>#${esc(tree.num_pt || tree.cod_alb)}</strong></div><div class="tree-result-grid">${details}</div>${expandButton}<p class="tree-data-source-note">${detailsNote}</p><div class="tree-result-actions"><a class="btn btn-primary tree-navigate" href="${navigationUrl}" target="_blank" rel="noopener">NAVIGA VERSO L’ALBERO</a><button class="btn tree-whazzup-share" type="button">INVIA TRAMITE WHAZZUP</button><button class="btn tree-street-view" type="button">🌐 VISTA 360° E PERCORSO</button><button class="btn tree-maintenance-open" type="button">🪚 MANUTENZIONE</button><button class="btn tree-work-order-open" type="button">✂️ CREA CANTIERE POTATURA / ABBATTIMENTO</button></div><section class="tree-maintenance-panel hidden" aria-live="polite"></section>`;
    const detailsToggle = result.querySelector(".tree-details-toggle");
    detailsToggle?.addEventListener("click", () => {
      const expanded = detailsToggle.getAttribute("aria-expanded") !== "true";
      result.querySelectorAll(".tree-detail-extra").forEach((item) => { item.hidden = !expanded; });
      detailsToggle.setAttribute("aria-expanded", String(expanded));
      detailsToggle.textContent = expanded
        ? "MOSTRA SOLO I PRIMI 6 DETTAGLI"
        : `MOSTRA TUTTI I DETTAGLI (${detailEntries.length})`;
    });
    result.querySelector(".tree-whazzup-share")?.addEventListener("click", () => {
      openTreeShareInWhazzup(buildTreeWhazzupMessage(detailEntries, navigationUrl));
    });
    const streetViewButton = result.querySelector(".tree-street-view");
    streetViewButton?.addEventListener("click", () => {
      openTreeStreetView(tree, point, streetViewButton);
    });
    const maintenanceButton = result.querySelector(".tree-maintenance-open");
    maintenanceButton?.addEventListener("click", () => openTreeMaintenance(tree, maintenanceButton));
    const workOrderButton = result.querySelector(".tree-work-order-open");
    workOrderButton?.addEventListener("click", () => {
      const workOrders = window.HeraTreeWorkOrders;
      if (!workOrders?.open) {
        setStatus("La scheda Potature Abbattimenti non è ancora pronta. Ricarica l’app e riprova.", "error");
        return;
      }
      workOrders.open({
        tree,
        details: detailEntries,
        municipality: municipality.value,
        point,
        navigationUrl
      });
    });
    result.classList.remove("hidden");
    const currentMap = initializeMap();
    if (!currentMap) return;
    if (marker) marker.remove();
    marker = L.marker([point.lat, point.lon]).addTo(currentMap).bindPopup(`<strong>${esc(tree.classe || "Albero")}</strong><br>Numero ${esc(tree.num_pt || tree.cod_alb)}`).openPopup();
    currentMap.setView([point.lat, point.lon], 19);
    setTimeout(() => { currentMap.invalidateSize(); loadVisibleTrees(); }, 50);
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
  mapPlantingFilter?.addEventListener("change", () => {
    lastViewportKey = "";
    loadVisibleTrees();
  });
  mapStyle?.addEventListener("change", () => applyMapStyle(mapStyle.value));
  $("tree-qr-close-btn")?.addEventListener("click", stopScanner);
  dialog?.addEventListener("close", stopScanner);
  $("tree-qr-file")?.addEventListener("change", (event) => event.target.files?.[0] && scanFile(event.target.files[0]));
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && mapFullscreen && !dialog?.open) setMapFullscreen(false);
  });
  window.HeraTreeSearch = Object.freeze({ open: openPage, close: closePage });
})();

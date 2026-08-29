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
  const dialog = $("tree-qr-dialog");
  const video = $("tree-qr-video");
  const qrStatus = $("tree-qr-status");
  let map = null;
  let marker = null;
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

  function showTree(tree) {
    const point = tree.geo_point_2d;
    if (!point || !Number.isFinite(point.lat) || !Number.isFinite(point.lon)) throw new Error("Albero trovato, ma senza coordinate utilizzabili.");
    result.innerHTML = `<div class="tree-result-title"><div><small>Comune di Bologna</small><h2>${esc(tree.classe || "Specie non disponibile")}</h2></div><strong>#${esc(tree.num_pt || tree.cod_alb)}</strong></div><div class="tree-result-grid"><div><span>Numero punto</span><strong>${esc(tree.num_pt)}</strong></div><div><span>Codice albero</span><strong>${esc(tree.cod_alb)}</strong></div><div><span>Altezza</span><strong>${esc(tree.cl_h)}</strong></div><div><span>Circonferenza</span><strong>${esc(tree.classe_circonferenza_diametro)}</strong></div><div><span>Quartiere</span><strong>${esc(tree.quartiere)}</strong></div><div><span>Albero di pregio</span><strong>${tree.pregio === "S" ? "Sì" : "No"}</strong></div></div><a class="btn btn-primary tree-navigate" href="https://www.google.com/maps/dir/?api=1&destination=${point.lat},${point.lon}" target="_blank" rel="noopener">NAVIGA VERSO L’ALBERO</a>`;
    result.classList.remove("hidden");
    mapNode.classList.remove("hidden");
    if (!map) {
      map = L.map(mapNode).setView([point.lat, point.lon], 18);
      L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 20, attribution: "&copy; OpenStreetMap" }).addTo(map);
    }
    if (marker) marker.remove();
    marker = L.marker([point.lat, point.lon]).addTo(map).bindPopup(`<strong>${esc(tree.classe || "Albero")}</strong><br>Numero ${esc(tree.num_pt || tree.cod_alb)}`).openPopup();
    map.setView([point.lat, point.lon], 19);
    setTimeout(() => map.invalidateSize(), 50);
  }

  form?.addEventListener("submit", async (event) => {
    event.preventDefault();
    setStatus("Ricerca nel censimento ufficiale…");
    result.classList.add("hidden"); mapNode.classList.add("hidden");
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
  $("tree-qr-close-btn")?.addEventListener("click", stopScanner);
  dialog?.addEventListener("close", stopScanner);
  $("tree-qr-file")?.addEventListener("change", (event) => event.target.files?.[0] && scanFile(event.target.files[0]));
})();

(() => {
  "use strict";

  const PAGE_ID = "verde-bologna-page";
  const STYLE_ID = "verde-bologna-operativo-style";
  const SELECT_ID = "verde-bologna-operativo-category";
  const MOBILE_QUERY = "(max-width: 760px)";
  const PARKS_DATASET_ID = "carta-tecnica-comunale-toponimi-parchi-e-giardini";
  const API_ROOT = "https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets";
  const DATASETS = [
    ["un_gest", "🌳 Aree verdi in manutenzione"],
    ["alberi-manutenzioni", "🌲 Alberi singoli"],
    ["popolazione-arborea", "🌴 Popolazioni arboree"],
    ["siepi", "🌿 Siepi in manutenzione"],
    ["attrezzature_ludiche_ginniche_sportive", "🛝 Giochi e attrezzature"],
    ["arredo", "🪑 Arredo urbano comunale"],
    ["sgambatura_cani", "🐕 Aree cani"],
    [PARKS_DATASET_ID, "🏞️ Parchi e giardini"],
    ["aree-verdi_entrate_centroidi", "🚪 Ingressi aree verdi"],
    ["aree-ortive", "🥕 Aree ortive"],
    ["verde_privato_urbanizzato", "🏡 Verde privato"]
  ];

  const $ = (id) => document.getElementById(id);
  let verdeMap = null;
  let codeLayer = null;
  let parkBoundaryLayer = null;
  let resultsObserver = null;
  let refreshTimer = 0;
  let mapCaptureInstalled = false;
  const parkBoundaryCache = new Map();

  function injectStyle() {
    if ($(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
      .verde-bologna-operativo-card{display:none}
      .verde-bologna-code-marker-wrap{background:transparent!important;border:0!important}
      .verde-bologna-code-marker{display:flex;align-items:center;justify-content:center;min-width:38px;height:28px;padding:0 7px;border:2px solid #fff;border-radius:15px;color:#fff;background:#08783f;box-shadow:0 2px 7px rgba(0,0,0,.42);font-size:.72rem;font-weight:900;white-space:nowrap;line-height:1}
      .verde-bologna-code-marker.is-park{background:#12623a}
      .verde-bologna-code-marker.is-fallback{background:#58697b}
      @media ${MOBILE_QUERY}{
        .verde-bologna-page{background:#f1f6fb!important;color:#10264a!important}
        .verde-bologna-shell{padding:0 10px 18px!important}
        .verde-bologna-header{position:sticky!important;top:0!important;z-index:1010!important;display:flex!important;grid-template-columns:none!important;gap:10px!important;align-items:center!important;margin:0 -10px 10px!important;padding:max(10px,env(safe-area-inset-top)) 10px 10px!important;background:#fff!important;border-bottom:1px solid #d9e3ef!important}
        .verde-bologna-header .btn{min-width:auto!important;padding:8px 10px!important}
        .verde-bologna-header h1{margin:0!important;font-size:1.18rem!important;color:#10264a!important}
        .verde-bologna-header p{margin:2px 0 0!important;font-size:.72rem!important;color:#55708f!important}
        .verde-bologna-badge{display:none!important}
        .verde-bologna-hero,.verde-bologna-section-title,.verde-bologna-datasets{display:none!important}
        .verde-bologna-operativo-card{display:grid;gap:9px;margin:10px 0;padding:14px;border-radius:18px;background:#fff;box-shadow:0 8px 25px rgba(26,55,91,.1)}
        .verde-bologna-operativo-card label{font-size:.82rem;font-weight:900;color:#10264a}
        .verde-bologna-operativo-card select{width:100%;min-height:48px;padding:10px 12px;border:1px solid #b9c9da;border-radius:10px;background:#fff;font:inherit;color:#10264a}
        .verde-bologna-operativo-hint{margin:0;color:#55708f;font-size:.78rem;line-height:1.4}
        .verde-bologna-browser{display:block!important;margin:10px 0 0!important;padding:14px!important;border:0!important;border-radius:18px!important;background:#fff!important;box-shadow:0 8px 25px rgba(26,55,91,.1)!important}
        .verde-bologna-browser.hidden{display:block!important}
        .verde-bologna-browser-head{display:flex!important;align-items:center!important;gap:8px!important}
        .verde-bologna-browser-head>div{min-width:0;flex:1}
        .verde-bologna-browser-head h2{font-size:1rem!important;color:#10264a!important;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
        .verde-bologna-browser-head p{display:none!important}
        .verde-bologna-source-link{width:auto!important;min-height:36px!important;padding:7px 9px!important;border-radius:9px!important;font-size:.69rem!important;white-space:nowrap!important}
        .verde-bologna-search{display:grid!important;grid-template-columns:1fr 1fr!important;gap:8px!important;margin:10px 0!important}
        .verde-bologna-search input{grid-column:1/-1!important;width:100%!important;min-height:48px!important;padding:10px 12px!important;border:1px solid #b9c9da!important;border-radius:10px!important;font-size:1rem!important}
        .verde-bologna-search .btn{min-height:44px!important;margin:0!important}
        .verde-bologna-status{margin:0 0 10px!important;padding:7px 10px!important;border-radius:9px!important;background:#edf4fb!important;color:#355777!important;font-size:.78rem!important}
        .verde-bologna-map-card{margin:0!important;padding:10px!important;border:0!important;border-radius:16px!important;background:#fff!important;box-shadow:none!important}
        .verde-bologna-map-toolbar{display:grid!important;grid-template-columns:1fr 1fr!important;gap:7px!important}
        .verde-bologna-map-toolbar strong{grid-column:1/-1!important;margin:0!important;font-size:.92rem!important}
        .verde-bologna-map-toolbar .btn{min-height:40px!important;padding:7px 8px!important;font-size:.72rem!important;font-weight:900!important;white-space:nowrap!important}
        .verde-bologna-map-status{margin:0!important;padding:6px 8px!important;border-radius:8px!important;background:#edf4fb!important;color:#355777!important;font-size:.72rem!important}
        .verde-bologna-map{height:52vh!important;min-height:360px!important;border-radius:12px!important;background:#e9eef4!important}
        .verde-bologna-code-marker{min-width:40px;height:29px;padding:0 7px;font-size:.72rem}
        .verde-bologna-results{gap:8px!important;margin-top:10px!important}
        .verde-bologna-result{grid-template-columns:1fr!important;gap:8px!important;padding:11px!important;border-radius:12px!important;background:#f8fbfd!important}
        .verde-bologna-result h3{font-size:.95rem!important;color:#10264a!important}
        .verde-bologna-result p{font-size:.78rem!important}
        .verde-bologna-result-actions{display:grid!important;grid-template-columns:1fr 1fr!important;gap:6px!important;justify-content:stretch!important}
        .verde-bologna-result-actions .btn,.verde-bologna-result-actions a{width:100%!important;min-height:40px!important;display:flex!important;align-items:center!important;justify-content:center!important;font-size:.72rem!important}
        .verde-bologna-details{grid-template-columns:1fr 1fr!important;gap:6px!important}
        .verde-bologna-details div{padding:7px!important}
        .verde-bologna-details span{font-size:.66rem!important}
        .verde-bologna-details strong{font-size:.75rem!important}
        .verde-bologna-load-more{min-height:44px!important;margin-top:9px!important}
        .verde-bologna-map-card.is-fullscreen{z-index:12060!important;padding:max(8px,env(safe-area-inset-top)) 8px max(8px,env(safe-area-inset-bottom))!important;border-radius:0!important;background:#f1f6fb!important}
        .verde-bologna-map-card.is-fullscreen .verde-bologna-map{height:100%!important;min-height:0!important}
      }
    `;
    document.head.appendChild(style);
  }

  function ensureOperationalCard(page) {
    if (!page || $(SELECT_ID)) return;
    const browser = $("verde-bologna-browser");
    if (!browser) return;
    const card = document.createElement("section");
    card.className = "verde-bologna-operativo-card";
    card.innerHTML = `
      <label for="${SELECT_ID}">Cosa devi cercare?</label>
      <select id="${SELECT_ID}" aria-label="Categoria Verde Bologna">
        ${DATASETS.map(([id,label]) => `<option value="${id}">${label}</option>`).join("")}
      </select>
      <p class="verde-bologna-operativo-hint">Scegli la categoria, cerca per nome/codice/via e usa la mappa come nel Catasto alberi.</p>`;
    browser.parentNode.insertBefore(card, browser);

    const select = $(SELECT_ID);
    select?.addEventListener("change", () => {
      const id = select.value;
      const button = page.querySelector(`[data-vb-open="${CSS.escape(id)}"]`);
      if (button) button.click();
      scheduleCodeMarkers();
    });
  }

  function syncFromCards(page) {
    const select = $(SELECT_ID);
    if (!select || !page) return;
    page.addEventListener("click", (event) => {
      const button = event.target?.closest?.("[data-vb-open]");
      const id = button?.getAttribute("data-vb-open");
      if (id && [...select.options].some((option) => option.value === id)) {
        select.value = id;
        parkBoundaryLayer?.clearLayers?.();
        scheduleCodeMarkers();
      }
    }, true);
  }

  function primeDefaultDataset(page) {
    if (!page || !window.matchMedia(MOBILE_QUERY).matches) return;
    const browser = $("verde-bologna-browser");
    if (browser && !browser.classList.contains("hidden")) return;
    page.querySelector('[data-vb-open="un_gest"]')?.click();
  }

  function normalizeKey(value) {
    return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "");
  }

  function detailsForCard(card) {
    const details = new Map();
    card?.querySelectorAll?.(".verde-bologna-details div").forEach((row) => {
      const label = normalizeKey(row.querySelector("span")?.textContent);
      const value = String(row.querySelector("strong")?.textContent || "").trim();
      if (label && value) details.set(label, value);
    });
    return details;
  }

  function usableCode(value) {
    const text = String(value || "").trim();
    if (!text || text === "—" || /^0(?:[.,]0+)?$/.test(text) || /^(null|undefined)$/i.test(text)) return "";
    return text;
  }

  function codeForCard(card) {
    const details = detailsForCard(card);
    const candidates = [
      "codvia", "codogg", "codice", "codiceoggetto", "codalb", "numpt",
      "idtopon", "progressiv", "objectid", "id", "inpatrim", "ref"
    ];
    for (const key of candidates) {
      const value = usableCode(details.get(key));
      if (value) return { value, fallback: !["codvia", "codogg", "codice", "codiceoggetto"].includes(key), key };
    }
    return null;
  }

  function titleForCard(card) {
    const details = detailsForCard(card);
    const isParks = activeDatasetId() === PARKS_DATASET_ID;
    const officialName = isParks ? (details.get("completo") || details.get("porzione")) : "";
    if (officialName) {
      const h3 = card?.querySelector?.("h3");
      if (h3 && h3.textContent !== officialName) h3.textContent = officialName;
      return officialName;
    }
    return String(card?.querySelector?.("h3")?.textContent || "Elemento verde").trim();
  }

  function centerFromCard(card) {
    const anchor = card?.querySelector?.('.verde-bologna-result-actions a[href*="destination="]');
    if (!anchor) return null;
    try {
      const url = new URL(anchor.href, window.location.href);
      const raw = url.searchParams.get("destination") || "";
      const [lat, lon] = raw.split(",").map(Number);
      return Number.isFinite(lat) && Number.isFinite(lon) ? { lat, lon } : null;
    } catch (_) { return null; }
  }

  function activeDatasetId() {
    const select = $(SELECT_ID);
    if (select?.value) return select.value;
    const source = $("verde-bologna-source-link")?.getAttribute("href") || "";
    return DATASETS.find(([id]) => source.includes(`/dataset/${id}/`))?.[0] || "un_gest";
  }

  function captureMapFactory() {
    if (mapCaptureInstalled || !window.L?.map) return Boolean(mapCaptureInstalled);
    const originalMap = window.L.map;
    if (originalMap.__verdeBolognaCapture) {
      mapCaptureInstalled = true;
      return true;
    }
    const wrappedMap = function(element, options) {
      const map = originalMap(element, options);
      const node = typeof element === "string" ? document.getElementById(element) : element;
      if (node?.id === "verde-bologna-map") {
        verdeMap = map;
        codeLayer = window.L.layerGroup().addTo(verdeMap);
        parkBoundaryLayer = window.L.layerGroup().addTo(verdeMap);
        scheduleCodeMarkers();
      }
      return map;
    };
    wrappedMap.__verdeBolognaCapture = true;
    wrappedMap.__originalMap = originalMap;
    window.L.map = wrappedMap;
    mapCaptureInstalled = true;
    return true;
  }

  function removeSimpleGreenPointMarkers() {
    if (!verdeMap || !window.L?.CircleMarker) return;
    const cleanGroup = (group) => {
      if (!group?.eachLayer) return;
      const removals = [];
      group.eachLayer((child) => {
        const fill = String(child?.options?.fillColor || "").toLowerCase();
        const radius = Number(child?.options?.radius);
        if (child instanceof window.L.CircleMarker && fill === "#18854b" && radius === 7) removals.push(child);
        else if (child?.eachLayer) cleanGroup(child);
      });
      removals.forEach((child) => group.removeLayer?.(child));
    };
    verdeMap.eachLayer((layer) => {
      if (layer !== codeLayer && layer !== parkBoundaryLayer && layer?.eachLayer) cleanGroup(layer);
    });
  }

  function scheduleCodeMarkers() {
    window.clearTimeout(refreshTimer);
    refreshTimer = window.setTimeout(renderCodeMarkers, 40);
  }

  function renderCodeMarkers() {
    if (!verdeMap || !window.L || !codeLayer) return;
    codeLayer.clearLayers();
    const cards = [...document.querySelectorAll("#verde-bologna-results .verde-bologna-result")];
    const isParks = activeDatasetId() === PARKS_DATASET_ID;
    let labels = 0;
    cards.forEach((card, index) => {
      const code = codeForCard(card);
      const center = centerFromCard(card);
      if (!code || !center) return;
      const title = titleForCard(card);
      const classNames = ["verde-bologna-code-marker", isParks ? "is-park" : "", code.fallback ? "is-fallback" : ""].filter(Boolean).join(" ");
      const marker = window.L.marker([center.lat, center.lon], {
        icon: window.L.divIcon({
          className: "verde-bologna-code-marker-wrap",
          html: `<span class="${classNames}">${String(code.value).replace(/[&<>"']/g, (char) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]))}</span>`,
          iconSize: null,
          iconAnchor: [20, 14]
        }),
        keyboard: true,
        riseOnHover: true,
        title: `${code.value} · ${title}`
      }).addTo(codeLayer);
      marker.bindPopup(`<strong>${String(code.value).replace(/[&<>"']/g, "")}</strong><br>${String(title).replace(/[&<>"']/g, "")}`);
      marker.on("click", () => {
        card.scrollIntoView({ behavior: "smooth", block: "nearest" });
        if (isParks) void showManagedParkBoundary(card);
      });
      labels += 1;
      if (index === 0 && isParks) titleForCard(card);
    });
    if (labels) removeSimpleGreenPointMarkers();
  }

  function meaningfulParkWords(name) {
    const stop = new Set(["parco", "giardino", "giardini", "area", "verde", "del", "della", "delle", "dei", "degli", "di", "da", "il", "lo", "la", "i", "gli", "le"]);
    return String(name || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[^A-Z0-9 ]+/g, " ").split(/\s+/).filter((word) => word.length > 2 && !stop.has(word.toLowerCase()));
  }

  function escapeSearch(value) {
    return String(value || "").replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  }

  async function fetchManagedParkRecords(name) {
    const words = meaningfulParkWords(name);
    if (!words.length) return [];
    const cacheKey = words.join(" ");
    if (parkBoundaryCache.has(cacheKey)) return parkBoundaryCache.get(cacheKey);

    const queryWords = words.slice(-3).join(" ");
    const fetchByTerm = async (term) => {
      const params = new URLSearchParams({ limit: "100", where: `search(\"${escapeSearch(term)}\")` });
      const response = await fetch(`${API_ROOT}/un_gest/records?${params}`, { headers: { Accept: "application/json" } });
      if (!response.ok) return [];
      const payload = await response.json();
      return Array.isArray(payload?.results) ? payload.results : [];
    };

    let records = await fetchByTerm(queryWords);
    if (!records.length && words.length > 1) records = await fetchByTerm(words[words.length - 1]);
    const normalizedWords = words.map(normalizeKey);
    records = records.filter((record) => {
      const haystack = normalizeKey([record?.nome, record?.nome_ug, record?.ubicazione].filter(Boolean).join(" "));
      return normalizedWords.every((word) => haystack.includes(word));
    });
    parkBoundaryCache.set(cacheKey, records);
    return records;
  }

  async function showManagedParkBoundary(card) {
    if (activeDatasetId() !== PARKS_DATASET_ID || !verdeMap || !parkBoundaryLayer || !window.L) return;
    const parkName = titleForCard(card);
    const status = $("verde-bologna-map-status");
    if (status) status.textContent = `Cerco i confini gestionali di ${parkName}…`;
    try {
      const records = await fetchManagedParkRecords(parkName);
      parkBoundaryLayer.clearLayers();
      const bounds = window.L.latLngBounds([]);
      records.forEach((record) => {
        const shape = record?.geo_shape?.type === "Feature" ? record.geo_shape : (record?.geo_shape ? { type: "Feature", geometry: record.geo_shape, properties: {} } : null);
        if (!shape) return;
        const layer = window.L.geoJSON(shape, { style: { color: "#0b6b3a", weight: 4, fillColor: "#45a96a", fillOpacity: 0.16 } }).addTo(parkBoundaryLayer);
        const layerBounds = layer.getBounds?.();
        if (layerBounds?.isValid?.()) bounds.extend(layerBounds);
      });
      if (bounds.isValid()) {
        verdeMap.fitBounds(bounds.pad(0.08), { animate: false, maxZoom: 18 });
        if (status) status.textContent = `${records.length} unità gestionali comunali evidenziate per ${parkName}. Il codice resta visibile sulla mappa.`;
      } else if (status) status.textContent = `Nessun confine gestionale abbinato automaticamente a ${parkName}. Il codice e la posizione ufficiale restano disponibili.`;
    } catch (_) {
      if (status) status.textContent = `Confini gestionali non disponibili in questo momento. Il codice e la posizione ufficiale restano disponibili.`;
    }
  }

  function observeResults() {
    const node = $("verde-bologna-results");
    if (!node || resultsObserver) return;
    resultsObserver = new MutationObserver(scheduleCodeMarkers);
    resultsObserver.observe(node, { childList: true, subtree: true, characterData: true });
    node.addEventListener("click", (event) => {
      const showButton = event.target?.closest?.("[data-vb-map-index]");
      if (!showButton || activeDatasetId() !== PARKS_DATASET_ID) return;
      const card = showButton.closest(".verde-bologna-result");
      if (card) window.setTimeout(() => void showManagedParkBoundary(card), 30);
    }, true);
    scheduleCodeMarkers();
  }

  function observePage(page) {
    if (!page || page.dataset.operativoObserved === "1") return;
    page.dataset.operativoObserved = "1";
    const observer = new MutationObserver(() => {
      if (!page.classList.contains("hidden")) {
        ensureOperationalCard(page);
        window.setTimeout(() => primeDefaultDataset(page), 40);
        window.setTimeout(scheduleCodeMarkers, 80);
      } else {
        parkBoundaryLayer?.clearLayers?.();
        codeLayer?.clearLayers?.();
      }
    });
    observer.observe(page, { attributes: true, attributeFilter: ["class", "aria-hidden"] });
  }

  function install() {
    injectStyle();

    let mapAttempts = 0;
    const mapTimer = window.setInterval(() => {
      mapAttempts += 1;
      if (captureMapFactory() || mapAttempts > 120) window.clearInterval(mapTimer);
    }, 100);

    const attach = () => {
      const page = $(PAGE_ID);
      if (!page) return false;
      ensureOperationalCard(page);
      syncFromCards(page);
      observePage(page);
      observeResults();
      if (!page.classList.contains("hidden")) primeDefaultDataset(page);
      return true;
    };
    if (attach()) return;
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      if (attach() || attempts > 80) window.clearInterval(timer);
    }, 100);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();
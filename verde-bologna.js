(() => {
  "use strict";

  const API_ROOT = "https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets";
  const PAGE_SIZE = 100;
  const CACHE_TTL_MS = 10 * 60 * 1000;
  const FETCH_TIMEOUT_MS = 10000;
  const VIEWPORT_DELAY_MS = 350;
  const VIEWPORT_MAX_RECORDS = 500;
  const MOBILE_VIEWPORT_MAX_RECORDS = 140;
  const VIEWPORT_LIST_LIMIT = 60;
  const MOBILE_FULL_GEOMETRY_LIMIT = 40;
  const MOBILE_LABEL_MARKER_LIMIT = 80;
  const MOBILE_QUERY = "(max-width: 760px)";
  const PAGE_ID = "verde-bologna-page";
  const STYLE_ID = "verde-bologna-style";
  const MENU_BUTTON_ID = "open-verde-bologna-btn";
  const CACHE_PREFIX = "varga-verde-bologna:";
  const TREE_RETURN_KEY = "varga-verde-bologna:return-from-tree";
  const MAP_CREATED_EVENT = "hera:verde-bologna-map-created";
  const CATEGORY_OPENED_EVENT = "hera:verde-bologna-category-opened";
  const CATEGORY_CLOSED_EVENT = "hera:verde-bologna-category-closed";

  const DATASETS = Object.freeze([
    { id: "un_gest", icon: "🌳", title: "Aree verdi in manutenzione", short: "Aiuole, parchi, giardini, verde scolastico, sportivo e stradale.", priority: true, titleFields: ["nome_ug", "nome", "ubicazione"], searchHint: "nome, via, quartiere o tipo di area" },
    { id: "alberi-manutenzioni", icon: "🌲", title: "Alberi singoli", short: "Catasto degli alberi con numero punto, specie, caratteristiche e coordinate.", delegate: "open-tree-search-btn" },
    { id: "popolazione-arborea", icon: "🌴", title: "Popolazioni arboree", short: "Gruppi e superfici con popolazioni arboree e arbustive.", titleFields: ["classe", "classe_popolamento", "in_patrim"], codeFields: ["in_patrim"], searchHint: "patrimonio, classe o tipologia" },
    { id: "siepi", icon: "🌿", title: "Siepi in manutenzione", short: "Specie, tipologia, lunghezza, altezza, larghezza e superficie di potatura.", titleFields: ["classe", "classe_tipo_siepe", "in_patrim"], codeFields: ["in_patrim"], searchHint: "patrimonio, specie o tipo di siepe" },
    { id: "attrezzature_ludiche_ginniche_sportive", icon: "🛝", title: "Giochi e attrezzature sportive", short: "Attrezzature ludiche, ginniche e sportive presenti sul territorio.", titleFields: ["classe", "categoria", "presenza"], searchHint: "categoria o tipo di attrezzatura" },
    { id: "arredo", icon: "🪑", title: "Arredo urbano comunale", short: "Arredi censiti dal Comune di Bologna, separati dai dati OSM.", titleFields: ["classe_arredo", "zona_prossimita", "quartiere"], searchHint: "tipo di arredo, zona o quartiere" },
    { id: "sgambatura_cani", icon: "🐕", title: "Aree cani", short: "Aree di sgambatura cani in manutenzione comunale.", titleFields: ["nome", "nomezona", "cod_ug"], codeFields: ["id", "cod_ug", "cod_pre"], searchHint: "nome, codice, zona o quartiere" },
    { id: "carta-tecnica-comunale-toponimi-parchi-e-giardini", icon: "🏞️", title: "Parchi e giardini", short: "Toponimi ufficiali dei parchi e giardini del Comune.", titleFields: ["nomevia", "tipo", "codvia"], codeFields: ["codvia"], searchHint: "NOMEVIA, CODVIA o quartiere" },
    { id: "aree-verdi_entrate_centroidi", icon: "🚪", title: "Ingressi aree verdi", short: "Centroidi e ingressi delle maggiori aree verdi, utili per la navigazione.", titleFields: ["nome", "ubicazione", "tipo_di_area"], searchHint: "nome, ubicazione, tipo o quartiere" },
    { id: "aree-ortive", icon: "🥕", title: "Aree ortive", short: "Orti comunali, gestori, indirizzi e informazioni disponibili.", titleFields: ["denominazione_orto", "indirizzo_orto", "gestore_orto"], geoField: "geopoint", searchHint: "nome dell’orto, indirizzo, gestore o quartiere" },
    { id: "verde_privato_urbanizzato", icon: "🏡", title: "Verde privato", short: "Verde privato nel territorio urbanizzato, mantenuto separato dal verde pubblico.", privateGreen: true, titleFields: ["codice_ogg", "patrimonio", "quartiere"], codeFields: ["codice_ogg"], searchHint: "codice oggetto, patrimonio o quartiere" }
  ]);

  const state = {
    datasetId: "un_gest", query: "", offset: 0, total: 0, records: [], map: null,
    baseLayer: null, hybridLabels: null, featureLayer: null, featureByIndex: new Map(), userMarker: null,
    userAccuracy: null, requestSerial: 0, fullscreen: false, requestAbort: null,
    viewportTimer: 0, viewportSerial: 0, viewportAbort: null, lastViewportKey: "", loadingViewport: false,
    mobileRenderer: null, categoryOpen: false, activationSerial: 0
  };

  const TILE_LAYERS = Object.freeze({
    classic: { url: "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "&copy; OpenStreetMap contributors" } },
    satellite: { url: "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "Tiles &copy; Esri" } },
    labels: { url: "https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, maxNativeZoom: 19, keepBuffer: 5, updateWhenZooming: false, updateWhenIdle: true, attribution: "Labels &copy; Esri", pane: "overlayPane" } }
  });

  const $ = (id) => document.getElementById(id);
  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[char]));

  function mobileView() {
    return window.matchMedia?.(MOBILE_QUERY)?.matches === true;
  }

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
      .verde-bologna-page{position:fixed;inset:0;z-index:1060;overflow:auto;-webkit-overflow-scrolling:touch;overscroll-behavior-y:contain;background:#eef5f0;color:#173426}
      .verde-bologna-page.hidden{display:none!important}.verde-bologna-shell{width:min(1180px,100%);margin:auto;padding:0 16px 28px}
      .verde-bologna-header{position:sticky;top:0;z-index:20;display:grid;grid-template-columns:auto minmax(0,1fr) auto;gap:14px;align-items:center;margin:0 -16px 18px;padding:max(12px,env(safe-area-inset-top)) max(16px,calc((100vw - 1180px)/2 + 16px)) 12px;background:rgba(255,255,255,.96);border-bottom:1px solid #cdded2;backdrop-filter:blur(14px)}
      .verde-bologna-header h1{margin:0;font-size:clamp(1.25rem,3vw,1.8rem);color:#154d2e}.verde-bologna-header p{margin:3px 0 0;color:#5f7868;font-size:.88rem}.verde-bologna-badge{padding:7px 10px;border-radius:999px;background:#e0f4e6;color:#146435;font-size:.76rem;font-weight:900;white-space:nowrap}
      .verde-bologna-hero{display:grid;grid-template-columns:minmax(0,1.25fr) minmax(280px,.75fr);gap:16px;margin-bottom:18px}.verde-bologna-hero-card,.verde-bologna-note{padding:18px;border:1px solid #cfe0d3;border-radius:20px;background:#fff;box-shadow:0 8px 24px rgba(31,78,47,.08)}
      .verde-bologna-hero-card h2,.verde-bologna-note h2{margin:0 0 7px;color:#174d30}.verde-bologna-hero-card p,.verde-bologna-note p{margin:0;color:#526d5b;line-height:1.5}.verde-bologna-note{background:#f7fbf8}.verde-bologna-note strong{color:#8a4b0f}
      .verde-bologna-section-title{display:flex;align-items:end;justify-content:space-between;gap:12px;margin:22px 0 10px}.verde-bologna-section-title h2{margin:0;color:#173f29}.verde-bologna-section-title span{color:#698072;font-size:.83rem}.verde-bologna-datasets{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}
      .verde-bologna-dataset{display:grid;grid-template-rows:auto auto 1fr auto auto;gap:8px;min-height:190px;padding:15px;border:1px solid #ccdcd0;border-radius:18px;background:#fff;text-align:left;box-shadow:0 7px 18px rgba(29,80,47,.07)}.verde-bologna-dataset.is-active{border:2px solid #18854b;background:#f3fbf6}.verde-bologna-dataset.is-private{border-color:#e2c68f;background:#fffaf0}
      .verde-bologna-dataset-icon{font-size:1.65rem}.verde-bologna-dataset h3{margin:0;color:#1e4d31;font-size:1rem}.verde-bologna-dataset p{margin:0;color:#5e7464;line-height:1.42;font-size:.88rem}.verde-bologna-dataset small{color:#74877a;font-size:.72rem;overflow-wrap:anywhere}.verde-bologna-dataset .btn{margin-top:4px;width:100%;min-height:40px}
      .verde-bologna-page.is-category-open .verde-bologna-hero,.verde-bologna-page.is-category-open .verde-bologna-section-title,.verde-bologna-page.is-category-open .verde-bologna-datasets{display:none!important}
      .verde-bologna-browser{margin-top:18px;padding:16px;border:1px solid #cdded2;border-radius:20px;background:#fff;box-shadow:0 8px 24px rgba(31,78,47,.08)}.verde-bologna-browser-head{display:flex;justify-content:space-between;gap:14px;align-items:flex-start}.verde-bologna-browser-head h2{margin:0;color:#164d2e}.verde-bologna-browser-head p{margin:5px 0 0;color:#667d6d}
      .verde-bologna-source-link{display:inline-flex;align-items:center;justify-content:center;min-height:42px;padding:9px 12px;border:1px solid #9ec8aa;border-radius:12px;color:#155e35;background:#f1faf4;font-weight:900;text-decoration:none;white-space:nowrap}.verde-bologna-search{display:grid;grid-template-columns:minmax(0,1fr) auto auto;gap:8px;margin:14px 0}.verde-bologna-search input{min-height:46px;padding:10px 12px;border:1px solid #aebfb2;border-radius:11px;font:inherit}
      .verde-bologna-status{margin:0 0 12px;padding:9px 11px;border-radius:10px;background:#edf7f0;color:#315b3e;font-size:.85rem}.verde-bologna-status.error{background:#fff0ef;color:#9b281f}.verde-bologna-status.warning{background:#fff7df;color:#7b5304}.verde-bologna-map-card{display:grid;gap:8px;margin:12px 0 16px;padding:12px;border:1px solid #cdded2;border-radius:16px;background:#f9fcfa}
      .verde-bologna-map-toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.verde-bologna-map-toolbar strong{margin-right:auto}.verde-bologna-map-toolbar label{font-size:.78rem;font-weight:800}.verde-bologna-map-toolbar select{min-height:40px;padding:7px 9px;border:1px solid #aebfb2;border-radius:9px;background:#fff;font:inherit}.verde-bologna-map{height:min(48vh,520px);min-height:330px;border-radius:12px;overflow:hidden}.verde-bologna-map.is-interactive{touch-action:none!important}.verde-bologna-map-status{margin:0;color:#5d7464;font-size:.8rem}body.verde-bologna-fullscreen-open{overflow:hidden}.verde-bologna-map-card.is-fullscreen{position:fixed;inset:0;z-index:12060;margin:0;padding:max(8px,env(safe-area-inset-top)) max(8px,env(safe-area-inset-right)) max(8px,env(safe-area-inset-bottom)) max(8px,env(safe-area-inset-left));border-radius:0;background:#eef5f0;grid-template-rows:auto auto minmax(0,1fr)}.verde-bologna-map-card.is-fullscreen .verde-bologna-map{height:100%;min-height:0}
      .verde-bologna-marker-wrap{background:transparent!important;border:0!important}.verde-bologna-marker{display:flex;align-items:center;justify-content:center;min-width:38px;height:28px;padding:0 7px;border:2px solid #fff;border-radius:15px;color:#fff;background:#08783f;box-shadow:0 2px 7px rgba(0,0,0,.42);font-size:.72rem;font-weight:900;white-space:nowrap}.verde-bologna-popup-open{margin-top:7px;padding:7px 9px;border:0;border-radius:7px;background:#126b40;color:#fff;font-weight:800;cursor:pointer}
      .verde-bologna-results{display:grid;gap:10px}.verde-bologna-result{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:12px;padding:13px;border:1px solid #d6e3d9;border-radius:14px;background:#fbfdfb}.verde-bologna-result h3{margin:0;color:#174d30;font-size:1rem}.verde-bologna-result p{margin:4px 0 0;color:#617667;font-size:.84rem;line-height:1.4}.verde-bologna-result-actions{display:flex;gap:6px;align-items:center;flex-wrap:wrap;justify-content:flex-end}.verde-bologna-result-actions .btn,.verde-bologna-result-actions a{min-height:38px;padding:7px 10px;font-size:.76rem;text-decoration:none}
      .verde-bologna-details{grid-column:1/-1;display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:7px}.verde-bologna-details div{padding:8px;border-radius:9px;background:#eff6f1}.verde-bologna-details span{display:block;color:#6c7f70;font-size:.7rem}.verde-bologna-details strong{display:block;margin-top:3px;color:#294d35;font-size:.8rem;overflow-wrap:anywhere}.verde-bologna-load-more{display:block;width:100%;margin-top:12px;min-height:46px}.verde-bologna-empty{padding:18px;text-align:center;color:#65796a}
      .verde-bologna-sheet{display:none;position:fixed;inset:0;z-index:13070;overflow:auto;padding:max(10px,env(safe-area-inset-top)) 10px max(18px,env(safe-area-inset-bottom));background:#f1f6fb}.verde-bologna-sheet.is-open{display:block}.verde-bologna-sheet-head{position:sticky;top:-10px;z-index:2;display:flex;align-items:center;gap:10px;margin:-10px -10px 10px;padding:max(10px,env(safe-area-inset-top)) 10px 10px;background:rgba(255,255,255,.97);border-bottom:1px solid #d9e3ef}.verde-bologna-sheet-head h2{min-width:0;flex:1;margin:0;color:#10264a;font-size:1.05rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.verde-bologna-sheet-source{margin:0 0 10px;padding:9px 11px;border-radius:10px;background:#e8f6ed;color:#1d5b37;font-size:.78rem}.verde-bologna-sheet-actions{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;margin-bottom:10px}.verde-bologna-sheet-actions .btn{display:flex;align-items:center;justify-content:center;min-height:46px;text-decoration:none;text-align:center}.verde-bologna-sheet-fields{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:7px}.verde-bologna-sheet-field{padding:10px;border:1px solid #dce6ef;border-radius:11px;background:#fff}.verde-bologna-sheet-field[hidden]{display:none}.verde-bologna-sheet-field span{display:block;margin-bottom:3px;color:#617990;font-size:.68rem;font-weight:900;text-transform:uppercase;letter-spacing:.03em}.verde-bologna-sheet-field strong{display:block;color:#203e59;font-size:.82rem;line-height:1.35;white-space:pre-wrap;overflow-wrap:anywhere}.verde-bologna-details-toggle{grid-column:1/-1;width:100%;min-height:44px;margin-top:3px;border-color:#bed0e2!important;background:#edf4fb!important;color:#214f7d!important;font-weight:900}body.verde-bologna-sheet-open{overflow:hidden}
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
          <div><h1 id="verde-bologna-page-title">🌳 Verde Bologna</h1><p id="verde-bologna-page-subtitle">Scegli una categoria: verrà caricata una sola mappa alla volta</p></div>
          <span class="verde-bologna-badge">COMUNE DI BOLOGNA OPEN DATA</span>
        </header>
        <section class="verde-bologna-hero">
          <div class="verde-bologna-hero-card"><h2>Un solo punto per il verde comunale</h2><p>Aree verdi, alberi, siepi, arredi, giochi, aree cani, parchi, ingressi, orti e verde privato vengono interrogati direttamente dai dataset pubblici del Comune di Bologna. I dati sono caricati solo quando apri una categoria.</p></div>
          <div class="verde-bologna-note"><h2>Priorità ai dati ufficiali</h2><p><strong>Per Bologna il Comune è la fonte principale.</strong> OpenStreetMap resta disponibile nelle altre sezioni dell’app come integrazione cartografica, ma qui non sostituisce il censimento comunale.</p></div>
        </section>
        <div class="verde-bologna-section-title"><h2>Scegli una categoria</h2><span>11 categorie ufficiali · una mappa alla volta</span></div>
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
            <div class="verde-bologna-map-toolbar"><button id="verde-bologna-location-btn" class="btn" type="button" aria-label="Centra sulla mia posizione" title="La mia posizione">⌖</button><button id="verde-bologna-fullscreen-btn" class="btn" type="button" aria-label="Apri la mappa a schermo intero" title="Schermo intero" aria-pressed="false">⛶</button><label for="verde-bologna-map-style">Vista mappa</label><select id="verde-bologna-map-style" aria-label="Vista mappa"><option value="classic">Classica</option><option value="satellite">Satellite</option><option value="hybrid">Ibrida</option></select></div>
            <p id="verde-bologna-map-status" class="verde-bologna-map-status">Aumenta lo zoom per visualizzare gli elementi ufficiali nella zona. Quelli già caricati restano visibili.</p>
            <div id="verde-bologna-map" class="verde-bologna-map"></div>
          </section>
          <section id="verde-bologna-results" class="verde-bologna-results"></section>
          <button id="verde-bologna-load-more" class="btn verde-bologna-load-more hidden" type="button">CARICA ALTRI 100</button>
        </section>
      </div>
      <section id="verde-bologna-detail-sheet" class="verde-bologna-sheet" aria-hidden="true" aria-labelledby="verde-bologna-detail-title"></section>`;
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
    node.querySelectorAll("[data-vb-open]").forEach((button) => {
      button.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        openDataset(button.dataset.vbOpen);
      });
    });
  }

  function updatePageHeader(dataset = null) {
    const back = $("verde-bologna-back-btn"), title = $("verde-bologna-page-title"), subtitle = $("verde-bologna-page-subtitle");
    if (back) back.textContent = dataset ? "← CATEGORIE" : "← HOME";
    if (title) title.textContent = dataset ? `${dataset.icon} ${dataset.title}` : "🌳 Verde Bologna";
    if (subtitle) subtitle.textContent = dataset ? "Mappa dedicata · sono caricati solo i dati di questa categoria" : "Scegli una categoria: verrà caricata una sola mappa alla volta";
  }

  function currentDataset() { return DATASETS.find((item) => item.id === state.datasetId) || DATASETS[0]; }
  function setStatus(message, type = "") { const node = $("verde-bologna-status"); if (!node) return; node.textContent = message; node.className = `verde-bologna-status ${type}`.trim(); }
  const FIELD_LABELS = Object.freeze({
    in_patrim: "Codice patrimonio", cod_ug: "Codice unità gestionale", cod_pre: "Codice area",
    codvia: "Codice via", nomevia: "Nome ufficiale", codice_ogg: "Codice oggetto",
    nome_ug: "Nome unità gestionale", area_ug: "Superficie unità gestionale",
    classe_arredo: "Tipo di arredo", classe_conservazione: "Stato di conservazione",
    classe_tipo_siepe: "Tipo di siepe", sup_pot: "Superficie di potatura",
    denominazione_orto: "Nome dell’orto", indirizzo_orto: "Indirizzo dell’orto",
    gestore_orto: "Gestore", numero_orti: "Numero orti", data_agg: "Data aggiornamento",
    geo_point_2d: "Coordinate", geopoint: "Coordinate", geo_shape: "Geometria"
  });
  function friendlyLabel(key) { return FIELD_LABELS[key] || String(key || "").replace(/^geo_/, "").replace(/_/g, " ").replace(/\b\w/g, (char) => char.toUpperCase()); }

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

  function detailEntries(record) {
    const dataset = currentDataset();
    const priority = [...(dataset.codeFields || []), ...(dataset.titleFields || []), "quartiere", "ubicazione", "indirizzo", "data_agg"];
    const entries = Object.entries(record || {})
      .filter(([key, value]) => !isGeometryValue(key, value) && value !== null && value !== undefined && value !== "" && typeof value !== "function")
      .map(([key, value]) => ({ key, label: friendlyLabel(key), value: displayValue(value) }))
      .filter((item) => item.value);
    return entries.sort((a, b) => {
      const aIndex = priority.indexOf(a.key), bIndex = priority.indexOf(b.key);
      if (aIndex >= 0 || bIndex >= 0) return (aIndex < 0 ? 999 : aIndex) - (bIndex < 0 ? 999 : bIndex);
      return a.label.localeCompare(b.label, "it", { sensitivity: "base" });
    });
  }

  function findField(record, candidates) {
    const entries = Object.entries(record || {});
    for (const candidate of candidates) { const exact = entries.find(([key]) => key.toLowerCase() === candidate); if (exact && displayValue(exact[1])) return displayValue(exact[1]); }
    for (const candidate of candidates) { if (String(candidate).length < 3) continue; const fuzzy = entries.find(([key]) => key.toLowerCase().includes(candidate)); if (fuzzy && displayValue(fuzzy[1])) return displayValue(fuzzy[1]); }
    return "";
  }

  function recordTitle(record, index) { return findField(record, [...(currentDataset().titleFields || []), "denominazione", "nome", "name", "descrizione", "desc", "localizzazione", "specie", "classe", "toponimo", "codice", "id"]) || `${currentDataset().title} · ${index + 1}`; }
  function recordSubtitle(record) { return [findField(record, ["via", "indirizzo", "localita", "quartiere"]), findField(record, ["tipo", "tipologia", "classe", "specie"])].filter(Boolean).filter((value, index, values) => values.indexOf(value) === index).join(" · "); }
  function recordCode(record, index) {
    const dataset = currentDataset();
    const official = findField(record, [...(dataset.codeFields || []), "codice_ogg", "codice", "cod_ug", "cod_pre", "codvia", "in_patrim", "objectid", "id"]);
    return official ? { value: official, official: true } : { value: String(index + 1), official: false };
  }

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

  function applyMapStyle(style) {
    if (!state.map) return;
    state.baseLayer?.remove();
    state.hybridLabels?.remove();
    state.hybridLabels = null;
    const selected = style === "satellite" || style === "hybrid" ? TILE_LAYERS.satellite : TILE_LAYERS.classic;
    state.baseLayer = L.tileLayer(selected.url, selected.options).addTo(state.map);
    if (style === "hybrid") state.hybridLabels = L.tileLayer(TILE_LAYERS.labels.url, TILE_LAYERS.labels.options).addTo(state.map);
  }

  function initializeMap() {
    if (state.map) { syncMapInteractionMode(); return state.map; }
    if (!window.L || !$("verde-bologna-map")) {
      setStatus("Mappa non disponibile in questo momento. Torna alle categorie e riprova.", "error");
      return null;
    }
    const createMap = () => {
      state.map = L.map($("verde-bologna-map"), { zoomControl: true, zoomAnimation: false, fadeAnimation: false, markerZoomAnimation: false }).setView([44.4949, 11.3426], 15);
      applyMapStyle($("verde-bologna-map-style")?.value || "classic");
      state.featureLayer = L.layerGroup().addTo(state.map);
      state.map.on("moveend zoomend", scheduleViewportLoad);
      syncMapInteractionMode();
      window.dispatchEvent(new CustomEvent(MAP_CREATED_EVENT, { detail: { map: state.map, datasetId: state.datasetId } }));
      return state.map;
    };
    try { return createMap(); }
    catch (error) {
      try { state.map?.remove?.(); } catch (_) {}
      state.map = null;
      state.baseLayer = null;
      state.hybridLabels = null;
      state.featureLayer = null;
      renewMapContainer();
      try { return createMap(); }
      catch (retryError) {
        try { state.map?.remove?.(); } catch (_) {}
        state.map = null;
        state.featureLayer = null;
        setStatus(`Impossibile inizializzare la mappa: ${retryError?.message || error?.message || "errore sconosciuto"}. Torna alle categorie e riprova.`, "error");
        return null;
      }
    }
  }

  function resizeMap() { requestAnimationFrame(() => state.map?.invalidateSize({ pan: false, animate: false })); setTimeout(() => state.map?.invalidateSize({ pan: false, animate: false }), 180); }

  function syncMapInteractionMode() {
    if (!state.map) return;
    $("verde-bologna-map")?.classList.toggle("is-interactive", mobileView() || state.fullscreen);
    state.map.dragging?.enable?.();
    state.map.touchZoom?.enable?.();
    state.map.doubleClickZoom?.enable?.();
  }

  function addGeometryToMap(record, index, combinedBounds, { lightweight = false } = {}) {
    const geometry = geometryOf(record); if (!geometry || !state.featureLayer) return null;
    const title = recordTitle(record, index), code = recordCode(record, index); let layer = null;
    try {
      const markerIcon = () => L.divIcon({ className: "verde-bologna-marker-wrap", html: `<span class="verde-bologna-marker${code.official ? "" : " is-map-number"}">${esc(code.value)}</span>`, iconSize: null, iconAnchor: [20, 14], popupAnchor: [0, -15] });
      const popupHtml = `<strong>${esc(title)}</strong>${code.official ? `<br>Codice: ${esc(code.value)}` : ""}<br><span>Comune di Bologna · fonte ufficiale</span><br><button type="button" class="verde-bologna-popup-open" data-vb-popup-index="${index}">APRI SCHEDA</button>`;
      const attachPopup = (target) => {
        target.bindPopup(popupHtml);
        target.on("popupopen", (event) => {
          event.popup?.getElement?.()?.querySelector?.("[data-vb-popup-index]")?.addEventListener("click", () => openDetailSheet(index), { once: true });
        });
      };
      const center = centerOfGeometry(geometry);
      if (lightweight && center) {
        const showLabel = state.records.length <= MOBILE_LABEL_MARKER_LIMIT;
        if (!state.mobileRenderer && L.canvas) state.mobileRenderer = L.canvas({ padding: 0.35 });
        layer = showLabel
          ? L.marker([center.lat, center.lon], { icon: markerIcon(), keyboard: true, riseOnHover: true, title: `${code.official ? code.value + " · " : ""}${title}` })
          : L.circleMarker([center.lat, center.lon], { renderer: state.mobileRenderer, radius: 6, color: "#08783f", weight: 2, fillColor: "#31b96b", fillOpacity: 0.82 });
        layer.addTo(state.featureLayer);
        attachPopup(layer);
        combinedBounds.extend([center.lat, center.lon]);
        state.featureByIndex.set(index, { layer, center, marker: layer });
        return center;
      }
      layer = L.geoJSON({ type: "Feature", geometry, properties: {} }, {
        style: { color: "#187443", weight: 2, fillColor: "#3ba868", fillOpacity: 0.2 },
        pointToLayer: (_feature, latlng) => L.marker(latlng, {
          icon: markerIcon(),
          keyboard: true, riseOnHover: true, title: `${code.official ? code.value + " · " : ""}${title}`
        })
      }).addTo(state.featureLayer);
      attachPopup(layer);
      const bounds = layer.getBounds?.(); if (bounds?.isValid?.()) combinedBounds.extend(bounds);
      let centerMarker = null;
      if (center) {
        combinedBounds.extend([center.lat, center.lon]);
        if (geometry.type !== "Point" && geometry.type !== "MultiPoint") {
          centerMarker = L.marker([center.lat, center.lon], { icon: markerIcon(), keyboard: true, riseOnHover: true, title: `${code.official ? code.value + " · " : ""}${title}` }).addTo(state.featureLayer);
          attachPopup(centerMarker);
        }
      }
      state.featureByIndex.set(index, { layer, center, marker: centerMarker }); return center;
    } catch (_) { return null; }
  }

  function renderMap({ fitMap = true } = {}) {
    initializeMap(); state.featureLayer?.clearLayers(); state.featureByIndex.clear();
    const combinedBounds = L.latLngBounds([]); let geocoded = 0;
    const lightweight = mobileView() && state.records.length > MOBILE_FULL_GEOMETRY_LIMIT;
    state.records.forEach((record, index) => { if (addGeometryToMap(record, index, combinedBounds, { lightweight })) geocoded += 1; });
    if (fitMap && combinedBounds.isValid()) state.map.fitBounds(combinedBounds.pad(0.08), { animate: false, maxZoom: 17 });
    const mapStatus = $("verde-bologna-map-status"); if (mapStatus) mapStatus.textContent = geocoded ? `${geocoded} elementi ufficiali visualizzati${lightweight ? " in modalità mobile leggera" : ""}. Tocca un elemento per aprire la scheda.` : "I record caricati non espongono coordinate utilizzabili in questa pagina; i dati testuali restano consultabili.";
    resizeMap();
  }

  function focusResult(index) {
    const item = state.featureByIndex.get(index); if (!item || !state.map) return;
    const bounds = item.layer?.getBounds?.(); if (bounds?.isValid?.()) state.map.fitBounds(bounds.pad(0.2), { animate: false, maxZoom: 18 }); else if (item.center) state.map.setView([item.center.lat, item.center.lon], 18, { animate: false });
    try { const layers = item.layer?.getLayers?.() || []; (item.marker || layers[0] || item.layer)?.openPopup?.(); } catch (_) {}
    $("verde-bologna-map-card")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function closeDetailSheet() {
    const sheet = $("verde-bologna-detail-sheet");
    sheet?.classList.remove("is-open");
    sheet?.setAttribute("aria-hidden", "true");
    document.body.classList.remove("verde-bologna-sheet-open");
  }

  function openStreetView(record, index, button) {
    const center = centerOfGeometry(geometryOf(record));
    if (!center) return;
    const api = window.HeraStreetViewCards;
    if (typeof api?.openForCoordinates !== "function") {
      window.alert("Vista panoramica 360° non disponibile in questo momento. Riprova tra qualche secondo.");
      return;
    }
    api.openForCoordinates({ lat: center.lat, lng: center.lon }, button, {
      targetLabel: currentDataset().title,
      modalTitle: `🌐 Vista 360° · ${recordTitle(record, index)}`
    });
  }

  function openDetailSheet(index) {
    const record = state.records[index], sheet = $("verde-bologna-detail-sheet");
    if (!record || !sheet) return;
    const dataset = currentDataset(), entries = detailEntries(record), title = recordTitle(record, index);
    const center = centerOfGeometry(geometryOf(record));
    const navigationUrl = center ? `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${center.lat},${center.lon}`)}` : "";
    sheet.innerHTML = `
      <header class="verde-bologna-sheet-head"><button class="btn" type="button" data-vb-close-sheet>← INDIETRO</button><h2 id="verde-bologna-detail-title">${esc(dataset.icon)} ${esc(title)}</h2></header>
      <p class="verde-bologna-sheet-source">Comune di Bologna · dataset ufficiale “${esc(dataset.title)}” · sono mostrati tutti i campi valorizzati disponibili.</p>
      <div class="verde-bologna-sheet-actions">
        ${navigationUrl ? `<a class="btn btn-primary" href="${esc(navigationUrl)}" target="_blank" rel="noopener">NAVIGA VERSO L’ELEMENTO</a><button class="btn" type="button" data-vb-street-view>🌐 VISTA 360° E PERCORSO</button>` : ""}
      </div>
      <section class="verde-bologna-sheet-fields">
        ${entries.map((entry, entryIndex) => `<article class="verde-bologna-sheet-field"${entryIndex >= 6 ? " hidden" : ""}><span>${esc(entry.label)}</span><strong>${esc(entry.value)}</strong></article>`).join("")}
        ${entries.length > 6 ? `<button class="btn verde-bologna-details-toggle" type="button" aria-expanded="false" data-vb-toggle-details>MOSTRA TUTTI I DETTAGLI (${entries.length})</button>` : ""}
      </section>`;
    sheet.classList.add("is-open");
    sheet.setAttribute("aria-hidden", "false");
    document.body.classList.add("verde-bologna-sheet-open");
    sheet.scrollTop = 0;
    sheet.querySelector("[data-vb-close-sheet]")?.addEventListener("click", closeDetailSheet);
    sheet.querySelector("[data-vb-street-view]")?.addEventListener("click", (event) => openStreetView(record, index, event.currentTarget));
    sheet.querySelector("[data-vb-toggle-details]")?.addEventListener("click", (event) => {
      const button = event.currentTarget, expanded = button.getAttribute("aria-expanded") !== "true";
      sheet.querySelectorAll(".verde-bologna-sheet-field").forEach((field, fieldIndex) => { field.hidden = !expanded && fieldIndex >= 6; });
      button.setAttribute("aria-expanded", String(expanded));
      button.textContent = expanded ? "MOSTRA SOLO I PRIMI 6 DETTAGLI" : `MOSTRA TUTTI I DETTAGLI (${entries.length})`;
    });
  }

  function renderResults({ fitMap = true } = {}) {
    const node = $("verde-bologna-results"); if (!node) return;
    if (!state.records.length) { node.innerHTML = `<p class="verde-bologna-empty">Nessun record trovato con i filtri attuali.</p>`; $("verde-bologna-load-more")?.classList.add("hidden"); renderMap({ fitMap }); return; }
    const visibleRecords = state.query ? state.records : state.records.slice(0, VIEWPORT_LIST_LIMIT);
    node.innerHTML = visibleRecords.map((record, index) => {
      const title = recordTitle(record, index), subtitle = recordSubtitle(record), entries = detailEntries(record).slice(0, 6), center = centerOfGeometry(geometryOf(record));
      const navHref = center ? `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${center.lat},${center.lon}`)}` : "";
      return `<article class="verde-bologna-result"><div><h3>${esc(title)}</h3>${subtitle ? `<p>${esc(subtitle)}</p>` : ""}</div><div class="verde-bologna-result-actions">${center ? `<button class="btn" type="button" data-vb-map-index="${index}">MOSTRA</button>` : ""}<button class="btn" type="button" data-vb-detail-index="${index}">APRI SCHEDA</button>${center ? `<a class="btn btn-primary" href="${esc(navHref)}" target="_blank" rel="noopener">NAVIGA</a>` : ""}</div><div class="verde-bologna-details">${entries.map((entry) => `<div><span>${esc(entry.label)}</span><strong>${esc(entry.value)}</strong></div>`).join("")}</div></article>`;
    }).join("") + (!state.query && state.records.length > visibleRecords.length ? `<p class="verde-bologna-empty">Mappa completa · elenco limitato ai primi ${VIEWPORT_LIST_LIMIT} elementi per mantenere l’app fluida. Ingrandisci la mappa o usa la ricerca per restringere la zona.</p>` : "");
    node.querySelectorAll("[data-vb-map-index]").forEach((button) => button.addEventListener("click", () => focusResult(Number(button.dataset.vbMapIndex))));
    node.querySelectorAll("[data-vb-detail-index]").forEach((button) => button.addEventListener("click", () => openDetailSheet(Number(button.dataset.vbDetailIndex))));
    const more = $("verde-bologna-load-more"); more?.classList.toggle("hidden", state.records.length >= state.total || state.records.length === 0); if (more && !more.classList.contains("hidden")) more.textContent = `CARICA ALTRI 100 · ${state.records.length} / ${state.total}`;
    renderMap({ fitMap });
  }

  function localMatches(record, query) {
    const needle = String(query || "").trim().toLocaleLowerCase("it-IT"); if (!needle) return true;
    return Object.entries(record || {}).some(([key, value]) => !isGeometryValue(key, value) && displayValue(value).toLocaleLowerCase("it-IT").includes(needle));
  }

  async function requestRecords(datasetId, offset, query, signal) {
    const attempt = async (includeSearch) => {
      const cached = readCache(datasetId, offset, query, includeSearch); if (cached) return { payload: cached, includeSearch };
      const response = await fetch(apiUrl(datasetId, offset, query, includeSearch), { headers: { Accept: "application/json" }, signal });
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
    state.requestAbort?.abort();
    const controller = typeof AbortController === "function" ? new AbortController() : null;
    state.requestAbort = controller;
    const timeout = controller ? window.setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS) : 0;
    if (!append) { state.offset = 0; state.records = []; state.total = 0; }
    setStatus(`${append ? "Carico altri record" : "Interrogo il dataset ufficiale"} “${dataset.title}”…`);
    const more = $("verde-bologna-load-more"); if (more) more.disabled = true;
    try {
      const response = await requestRecords(dataset.id, offset, state.query, controller?.signal); if (serial !== state.requestSerial) return;
      const results = Array.isArray(response.payload?.results) ? response.payload.results : [];
      state.records = append ? [...state.records, ...results] : results; state.total = Number(response.payload?.total_count ?? state.records.length) || state.records.length; state.offset = state.records.length;
      renderResults(); const fallbackNote = response.localFallback ? " Ricerca applicata ai 100 record della pagina perché il filtro testuale remoto non era disponibile." : "";
      setStatus(`${state.records.length} record caricati${state.total ? ` su ${state.total}` : ""}.${fallbackNote}`, response.localFallback ? "warning" : "");
    } catch (error) { if (serial !== state.requestSerial) return; setStatus(error?.name === "AbortError" ? "Il dataset comunale sta impiegando troppo tempo. Riprova o aumenta lo zoom sulla mappa." : (error?.message || "Impossibile leggere il dataset del Comune di Bologna."), "error"); }
    finally { if (timeout) window.clearTimeout(timeout); if (state.requestAbort === controller) state.requestAbort = null; if (more) more.disabled = false; }
  }

  function distanceMeters(a, b) {
    const toRad = (value) => value * Math.PI / 180;
    const dLat = toRad(b.lat - a.lat), dLng = toRad(b.lng - a.lng);
    const value = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(a.lat)) * Math.cos(toRad(b.lat)) * Math.sin(dLng / 2) ** 2;
    return Math.round(6371000 * 2 * Math.atan2(Math.sqrt(value), Math.sqrt(1 - value)));
  }

  function scheduleViewportLoad() {
    window.clearTimeout(state.viewportTimer);
    state.viewportTimer = window.setTimeout(() => loadViewportRecords(), VIEWPORT_DELAY_MS);
  }

  async function loadViewportRecords({ force = false } = {}) {
    const page = $(PAGE_ID), dataset = currentDataset();
    if (!state.map || page?.classList.contains("hidden") || dataset.delegate || state.query) return;
    if (dataset.id === "carta-tecnica-comunale-toponimi-parchi-e-giardini") return;
    const zoom = state.map.getZoom();
    if (zoom < 15) {
      const mapStatus = $("verde-bologna-map-status");
      if (mapStatus) mapStatus.textContent = "Aumenta lo zoom almeno al livello 15. Gli elementi già caricati restano visibili.";
      return;
    }
    const center = state.map.getCenter();
    const viewportKey = `${dataset.id}:${zoom}:${center.lat.toFixed(4)}:${center.lng.toFixed(4)}`;
    if (!force && viewportKey === state.lastViewportKey) return;
    const radius = Math.min(2200, Math.max(80, distanceMeters(center, state.map.getBounds().getNorthEast()) + 40));
    const geoField = dataset.geoField || "geo_point_2d";
    const where = `within_distance(${geoField}, geom'POINT(${center.lng} ${center.lat})', ${radius}m)`;
    const firstParams = new URLSearchParams({ where, limit: String(PAGE_SIZE) });
    const firstUrl = `${API_ROOT}/${encodeURIComponent(dataset.id)}/records?${firstParams}`;
    const requestId = ++state.viewportSerial;
    state.viewportAbort?.abort();
    const controller = typeof AbortController === "function" ? new AbortController() : null;
    state.viewportAbort = controller;
    const timeout = controller ? window.setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS) : 0;
    state.loadingViewport = true;
    const mapStatus = $("verde-bologna-map-status");
    if (mapStatus) mapStatus.textContent = `Aggiorno ${dataset.title.toLocaleLowerCase("it-IT")} nella zona… Gli elementi attuali restano visibili.`;
    try {
      const firstResponse = await fetch(firstUrl, { headers: { Accept: "application/json" }, signal: controller?.signal });
      if (!firstResponse.ok) throw new Error(`Servizio comunale non disponibile (${firstResponse.status}).`);
      const first = await firstResponse.json();
      if (requestId !== state.viewportSerial) return;
      const total = Number(first.total_count) || 0;
      const viewportLimit = mobileView() ? MOBILE_VIEWPORT_MAX_RECORDS : VIEWPORT_MAX_RECORDS;
      if (total > viewportLimit) {
        const records = [...(first.results || [])].slice(0, PAGE_SIZE);
        state.records = records;
        state.total = total;
        state.offset = records.length;
        state.lastViewportKey = viewportKey;
        renderResults({ fitMap: false });
        setStatus(`${records.length} elementi mostrati su ${total} nella zona. Aumenta lo zoom per restringere e visualizzare tutti i risultati.`, "warning");
        if (mapStatus) mapStatus.textContent = `${records.length} elementi mostrati su ${total}. La mappa resta leggera; aumenta lo zoom per restringere la zona.`;
        return;
      }
      const records = [...(first.results || [])];
      for (let offset = PAGE_SIZE; offset < total; offset += PAGE_SIZE) {
        const response = await fetch(`${firstUrl}&offset=${offset}`, { headers: { Accept: "application/json" }, signal: controller?.signal });
        if (!response.ok) throw new Error(`Servizio comunale non disponibile (${response.status}).`);
        const payload = await response.json();
        if (requestId !== state.viewportSerial) return;
        records.push(...(payload.results || []));
      }
      state.records = records;
      state.total = total;
      state.offset = records.length;
      state.lastViewportKey = viewportKey;
      renderResults({ fitMap: false });
      setStatus(`${records.length} elementi ufficiali nella zona visibile. Cerca anche per ${dataset.searchHint || "nome o codice"}.`);
      if (mapStatus) mapStatus.textContent = `${records.length} elementi visualizzati. Tocca un codice o un numero per aprire la scheda completa.`;
    } catch (error) {
      if (error?.name === "AbortError" || requestId !== state.viewportSerial) return;
      if (mapStatus) mapStatus.textContent = `${error?.message || "Impossibile aggiornare la zona."} Gli elementi precedenti restano disponibili.`;
    } finally {
      if (timeout) window.clearTimeout(timeout);
      if (state.viewportAbort === controller) state.viewportAbort = null;
      if (requestId === state.viewportSerial) state.loadingViewport = false;
    }
  }

  function stopCategoryWork() {
    window.clearTimeout(state.viewportTimer);
    state.viewportTimer = 0;
    state.requestSerial += 1;
    state.viewportSerial += 1;
    state.requestAbort?.abort();
    state.viewportAbort?.abort();
    state.requestAbort = null;
    state.viewportAbort = null;
    state.loadingViewport = false;
  }

  function resetCategoryData() {
    stopCategoryWork();
    closeDetailSheet();
    if (state.fullscreen) setFullscreen(false);
    state.featureLayer?.clearLayers?.();
    state.featureByIndex.clear();
    state.records = [];
    state.offset = 0;
    state.total = 0;
    state.query = "";
    state.lastViewportKey = "";
    const results = $("verde-bologna-results");
    if (results) results.innerHTML = "";
    $("verde-bologna-load-more")?.classList.add("hidden");
  }

  function renewMapContainer() {
    const mapNode = $("verde-bologna-map");
    if (!mapNode?.parentNode) return mapNode;
    const freshMapNode = mapNode.cloneNode(false);
    freshMapNode.className = "verde-bologna-map";
    freshMapNode.removeAttribute("style");
    freshMapNode.removeAttribute("tabindex");
    freshMapNode.removeAttribute("aria-label");
    mapNode.replaceWith(freshMapNode);
    return freshMapNode;
  }

  function destroyCategoryMap() {
    resetCategoryData();
    const oldMap = state.map;
    try { oldMap?.remove?.(); } catch (_) { try { oldMap?.off?.(); } catch (_) {} }
    if (oldMap) window.dispatchEvent(new CustomEvent("hera:verde-bologna-map-destroyed", { detail: { map: oldMap } }));
    state.map = null;
    state.baseLayer = null;
    state.hybridLabels = null;
    state.featureLayer = null;
    state.userMarker = null;
    state.userAccuracy = null;
    state.mobileRenderer = null;
    renewMapContainer();
  }

  function showCategoryHub({ scroll = true } = {}) {
    state.activationSerial += 1;
    resetCategoryData();
    state.categoryOpen = false;
    const page = $(PAGE_ID);
    page?.classList.remove("is-category-open");
    $("verde-bologna-browser")?.classList.add("hidden");
    updatePageHeader();
    renderDatasetCards();
    window.dispatchEvent(new CustomEvent(CATEGORY_CLOSED_EVENT, { detail: { map: state.map } }));
    if (scroll) page?.scrollTo?.({ top: 0, behavior: "smooth" });
  }

  function openDataset(datasetId) {
    const dataset = DATASETS.find((item) => item.id === datasetId); if (!dataset) return;
    if (dataset.delegate) {
      state.activationSerial += 1;
      destroyCategoryMap();
      state.categoryOpen = false;
      const page = $(PAGE_ID), target = $(dataset.delegate);
      try { sessionStorage.setItem(TREE_RETURN_KEY, "1"); } catch (_) {}
      page?.classList.add("hidden"); page?.classList.remove("is-category-open"); page?.setAttribute("aria-hidden", "true");
      if (target) target.click();
      else { try { sessionStorage.removeItem(TREE_RETURN_KEY); } catch (_) {} openPage(); window.alert("Catasto alberi non disponibile in questo momento."); }
      return;
    }
    const activation = ++state.activationSerial;
    resetCategoryData();
    state.categoryOpen = true;
    state.datasetId = dataset.id;
    const page = $(PAGE_ID); page?.classList.add("is-category-open");
    $("verde-bologna-browser")?.classList.remove("hidden");
    const categorySelect = $("verde-bologna-operativo-category"); if (categorySelect) categorySelect.value = dataset.id;
    updatePageHeader(dataset);
    const title = $("verde-bologna-active-title"), description = $("verde-bologna-active-description"), source = $("verde-bologna-source-link"), query = $("verde-bologna-query");
    if (title) title.textContent = `${dataset.icon} ${dataset.title}`; if (description) description.textContent = dataset.short; if (source) source.href = sourcePageUrl(dataset.id); if (query) { query.value = ""; query.placeholder = `Cerca per ${dataset.searchHint || "nome, codice o via"}…`; }
    const results = $("verde-bologna-results"); if (results) results.innerHTML = `<p class="verde-bologna-empty">Sposta o ingrandisci la mappa per vedere gli elementi nella zona, oppure usa la ricerca.</p>`;
    setStatus(`Categoria “${dataset.title}” selezionata. Cerca per ${dataset.searchHint || "nome, codice o via"} oppure aumenta lo zoom sulla mappa.`);
    page?.scrollTo?.({ top: 0, behavior: "auto" });
    requestAnimationFrame(() => {
      if (activation !== state.activationSerial || !state.categoryOpen || state.datasetId !== dataset.id || page?.classList.contains("hidden")) return;
      const map = initializeMap();
      if (!map) return;
      resizeMap();
      window.dispatchEvent(new CustomEvent(CATEGORY_OPENED_EVENT, { detail: { datasetId: dataset.id, map } }));
      window.setTimeout(() => {
        if (activation !== state.activationSerial || !state.categoryOpen || state.datasetId !== dataset.id) return;
        if (dataset.id === "carta-tecnica-comunale-toponimi-parchi-e-giardini") loadRecords(); else loadViewportRecords({ force: true });
      }, 0);
    });
  }

  function openPage() {
    $("menu-close-btn")?.click(); $("home-page")?.classList.add("hidden"); const page = buildPage(); page.classList.remove("hidden"); page.setAttribute("aria-hidden", "false"); showCategoryHub({ scroll: false });
  }

  function closePage() { state.activationSerial += 1; destroyCategoryMap(); state.categoryOpen = false; const page = $(PAGE_ID); page?.classList.add("hidden"); page?.classList.remove("is-category-open"); page?.setAttribute("aria-hidden", "true"); $("home-page")?.classList.remove("hidden"); }

  function setFullscreen(active) {
    state.fullscreen = Boolean(active); $("verde-bologna-map-card")?.classList.toggle("is-fullscreen", state.fullscreen); document.body.classList.toggle("verde-bologna-fullscreen-open", state.fullscreen);
    const button = $("verde-bologna-fullscreen-btn"); if (button) { button.setAttribute("aria-pressed", String(state.fullscreen)); button.setAttribute("aria-label", state.fullscreen ? "Chiudi la mappa a schermo intero" : "Apri la mappa a schermo intero"); button.title = state.fullscreen ? "Chiudi mappa" : "Schermo intero"; button.textContent = state.fullscreen ? "✕" : "⛶"; } resizeMap();
    syncMapInteractionMode();
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
    $("verde-bologna-back-btn")?.addEventListener("click", () => { if (state.categoryOpen) showCategoryHub(); else closePage(); });
    $("verde-bologna-search-form")?.addEventListener("submit", (event) => { event.preventDefault(); state.query = $("verde-bologna-query")?.value.trim() || ""; if (state.query) loadRecords(); else loadViewportRecords({ force: true }); });
    $("verde-bologna-clear-btn")?.addEventListener("click", () => { state.query = ""; state.lastViewportKey = ""; if ($("verde-bologna-query")) $("verde-bologna-query").value = ""; loadViewportRecords({ force: true }); });
    $("verde-bologna-load-more")?.addEventListener("click", () => loadRecords({ append: true })); $("verde-bologna-location-btn")?.addEventListener("click", showUserLocation); $("verde-bologna-fullscreen-btn")?.addEventListener("click", () => setFullscreen(!state.fullscreen));
    $("verde-bologna-map-style")?.addEventListener("change", (event) => applyMapStyle(event.currentTarget.value));
    window.addEventListener("resize", syncMapInteractionMode, { passive: true });
    document.addEventListener("keydown", (event) => { if (event.key !== "Escape") return; if ($("verde-bologna-detail-sheet")?.classList.contains("is-open")) closeDetailSheet(); else if (state.fullscreen) setFullscreen(false); else if (!$(PAGE_ID)?.classList.contains("hidden")) { if (state.categoryOpen) showCategoryHub(); else closePage(); } });
  }

  function install() {
    injectStyle(); buildPage(); renderDatasetCards(); addMenuButton(); installEvents();
    window.HeraVerdeBologna = Object.freeze({ open: openPage, close: closePage, showCategories: showCategoryHub, openDataset, datasets: DATASETS.map(({ id, title }) => ({ id, title })) });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true }); else install();
})();

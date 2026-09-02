(() => {
  "use strict";

  const PAGE_ID = "verde-levato-page";
  const MENU_BUTTON_ID = "open-verde-levato-btn";
  const STYLE_ID = "verde-levato-style";
  const RECORDS_COLLECTION = "verdeLevatoRecords";
  const COMMESSE_COLLECTION = "verdeLevatoCommesse";
  const CONFIG_COLLECTION = "verdeLevatoConfig";
  const CONFIG_DOCUMENT = "access";
  const DEFAULT_CENTER = [44.4949, 11.3426];
  const CATEGORIES = Object.freeze([
    { id: "cantiere", icon: "🛠️", title: "Cantieri", description: "Aree di lavoro inserite manualmente dall’operatore autorizzato." },
    { id: "albero", icon: "🌳", title: "Alberi censiti", description: "Alberi singoli con posizione, specie, misure e stato vegetativo." },
    { id: "siepe", icon: "🌿", title: "Siepi", description: "Siepi censite con posizione, specie e dimensioni rilevate." }
  ]);

  const state = {
    records: [],
    commesse: [],
    commesseLoaded: false,
    commessePromise: null,
    adminEmails: [],
    category: "",
    query: "",
    canManage: false,
    globalAdmin: false,
    map: null,
    baseLayer: null,
    labelsLayer: null,
    markersLayer: null,
    locationMarker: null,
    locationAccuracy: null,
    loading: false
  };

  const TILE_LAYERS = Object.freeze({
    classic: { url: "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", options: { maxZoom: 20, maxNativeZoom: 19, attribution: "&copy; OpenStreetMap contributors" } },
    satellite: { url: "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, maxNativeZoom: 19, attribution: "Tiles &copy; Esri" } },
    labels: { url: "https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}", options: { maxZoom: 20, maxNativeZoom: 19, attribution: "Labels &copy; Esri", pane: "overlayPane" } }
  });

  const $ = (id) => document.getElementById(id);
  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  }[char]));
  const normalizeEmail = (value) => String(value || "").trim().toLocaleLowerCase("it-IT");
  const text = (value) => String(value ?? "").trim();
  const numberValue = (value) => {
    const parsed = Number(String(value ?? "").trim().replace(",", "."));
    return Number.isFinite(parsed) ? parsed : null;
  };

  function currentUser() {
    try { return auth?.currentUser || null; } catch (_) { return null; }
  }

  function store() {
    try { return db?.collection ? db : null; } catch (_) { return null; }
  }

  function firestoreTimestamp() {
    try { return firebase.firestore.FieldValue.serverTimestamp(); } catch (_) { return new Date(); }
  }

  function isGlobalAdmin() {
    try { return typeof canManageData === "function" && canManageData(); } catch (_) { return false; }
  }

  function injectStyle() {
    if ($(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
      .verde-levato-page{position:fixed;inset:0;z-index:1060;overflow:auto;-webkit-overflow-scrolling:touch;background:#eef5f0;color:#173426}
      .verde-levato-page.hidden,.verde-levato-hidden{display:none!important}.verde-levato-shell{width:min(1180px,100%);margin:auto;padding:0 14px 28px}
      .verde-levato-header{position:sticky;top:0;z-index:30;display:flex;align-items:center;gap:10px;margin:0 -14px 14px;padding:max(10px,env(safe-area-inset-top)) 14px 10px;background:rgba(255,255,255,.97);border-bottom:1px solid #cdded2;backdrop-filter:blur(14px)}
      .verde-levato-header-copy{min-width:0;flex:1}.verde-levato-header h1{margin:0;color:#154d2e;font-size:clamp(1.12rem,4vw,1.65rem)}.verde-levato-header p{margin:3px 0 0;color:#5f7868;font-size:.78rem}.verde-levato-badge{padding:6px 9px;border-radius:999px;background:#e0f4e6;color:#146435;font-size:.68rem;font-weight:900;white-space:nowrap}
      .verde-levato-hero{display:grid;grid-template-columns:minmax(0,1.3fr) minmax(250px,.7fr);gap:12px;margin-bottom:14px}.verde-levato-card{padding:15px;border:1px solid #cfe0d3;border-radius:18px;background:#fff;box-shadow:0 7px 20px rgba(31,78,47,.08)}
      .verde-levato-card h2,.verde-levato-card h3{margin:0 0 7px;color:#174d30}.verde-levato-card p{margin:0;color:#526d5b;line-height:1.45}.verde-levato-hero-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:12px}
      .verde-levato-access{display:grid;gap:9px}.verde-levato-admin-list{display:flex;gap:6px;flex-wrap:wrap}.verde-levato-admin-chip{padding:6px 9px;border-radius:999px;background:#edf4fb;color:#214f7d;font-size:.72rem;font-weight:800}
      .verde-levato-categories{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:11px}.verde-levato-category{display:grid;gap:7px;min-height:164px;padding:14px;border:1px solid #ccdcd0;border-radius:17px;background:#fff;text-align:left;box-shadow:0 6px 17px rgba(29,80,47,.07)}.verde-levato-category-icon{font-size:1.55rem}.verde-levato-category h3{margin:0;color:#1e4d31}.verde-levato-category p{margin:0;color:#5e7464;font-size:.82rem;line-height:1.35}.verde-levato-category strong{color:#12623a}
      .verde-levato-browser{display:grid;gap:10px}.verde-levato-browser.hidden{display:none!important}.verde-levato-toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.verde-levato-toolbar h2{min-width:180px;flex:1;margin:0;color:#164d2e}.verde-levato-search{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:7px}.verde-levato-search input{min-height:44px;padding:9px 11px;border:1px solid #aebfb2;border-radius:10px;font:inherit}
      .verde-levato-status{margin:0;padding:8px 10px;border-radius:9px;background:#edf7f0;color:#315b3e;font-size:.8rem}.verde-levato-status.error{background:#fff0ef;color:#9b281f}.verde-levato-status.warning{background:#fff7df;color:#7b5304}
      .verde-levato-map-card{position:relative;padding:10px;border:1px solid #cdded2;border-radius:15px;background:#f9fcfa}.verde-levato-map-toolbar{display:flex;gap:7px;align-items:center;margin-bottom:8px;flex-wrap:wrap}.verde-levato-map-toolbar strong{margin-right:auto}.verde-levato-map-toolbar select{min-height:38px;padding:6px 8px;border:1px solid #aebfb2;border-radius:9px;background:#fff}.verde-levato-map{height:min(48vh,510px);min-height:330px;border-radius:11px;overflow:hidden;touch-action:none}.verde-levato-marker-wrap{background:transparent!important;border:0!important}.verde-levato-marker{display:grid;place-items:center;width:34px;height:34px;border:2px solid #fff;border-radius:50%;background:#08783f;color:#fff;box-shadow:0 2px 7px rgba(0,0,0,.38);font-size:1rem}.verde-levato-marker.is-albero{background:#19713c}.verde-levato-marker.is-siepe{background:#65a30d}.verde-levato-marker.is-cantiere{background:#b45309}
      .verde-levato-results{display:grid;gap:9px}.verde-levato-result{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:10px;padding:12px;border:1px solid #d6e3d9;border-radius:13px;background:#fbfdfb}.verde-levato-result h3{margin:0;color:#174d30}.verde-levato-result p{margin:4px 0 0;color:#617667;font-size:.8rem}.verde-levato-result-actions{display:flex;gap:6px;align-items:center;flex-wrap:wrap}.verde-levato-result-details{grid-column:1/-1;display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:6px}.verde-levato-result-details div{padding:7px;border-radius:8px;background:#eff6f1}.verde-levato-result-details span{display:block;color:#6c7f70;font-size:.66rem}.verde-levato-result-details strong{display:block;margin-top:2px;color:#294d35;font-size:.76rem;overflow-wrap:anywhere}.verde-levato-empty{padding:17px;text-align:center;color:#65796a}
      .verde-levato-modal{position:fixed;inset:0;z-index:14000;overflow:auto;padding:max(10px,env(safe-area-inset-top)) 10px max(18px,env(safe-area-inset-bottom));background:#eef4f0}.verde-levato-modal.hidden{display:none!important}.verde-levato-modal-card{width:min(820px,100%);margin:auto;padding:14px;border-radius:18px;background:#fff;box-shadow:0 14px 38px rgba(15,50,28,.2)}.verde-levato-modal-head{display:flex;align-items:center;gap:8px;margin-bottom:12px}.verde-levato-modal-head h2{min-width:0;flex:1;margin:0;color:#174d30}.verde-levato-form{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.verde-levato-form label,.verde-levato-form fieldset{display:grid;gap:5px;margin:0;padding:0;border:0;color:#294d35;font-size:.78rem;font-weight:800}.verde-levato-form input,.verde-levato-form select,.verde-levato-form textarea{width:100%;min-height:43px;padding:9px 10px;border:1px solid #afc2b4;border-radius:9px;background:#fff;font:inherit;color:#173426}.verde-levato-form textarea{min-height:80px;resize:vertical}.verde-levato-form-wide{grid-column:1/-1}.verde-levato-location-box{grid-column:1/-1;display:grid;gap:8px;padding:11px;border:1px solid #b9d3c0;border-radius:12px;background:#f3faf5}.verde-levato-location-actions{display:flex;gap:7px;align-items:center;flex-wrap:wrap}.verde-levato-location-status{margin:0;color:#4d6c56;font-size:.76rem}.verde-levato-specific{display:contents}.verde-levato-form-footer{grid-column:1/-1;display:flex;justify-content:flex-end;gap:8px;margin-top:5px}.verde-levato-form-feedback{grid-column:1/-1;margin:0;color:#315b3e;font-size:.8rem}.verde-levato-form-feedback.error{color:#9b281f}
      .verde-levato-admin-form{display:grid;gap:10px}.verde-levato-admin-form input{min-height:44px;padding:9px 11px;border:1px solid #afc2b4;border-radius:9px;font:inherit}.verde-levato-admin-note{padding:9px;border-radius:9px;background:#fff7df;color:#72500b;font-size:.78rem}
      .verde-levato-commessa-box{grid-column:1/-1;display:grid;gap:8px;padding:11px;border:1px solid #b9d3c0;border-radius:12px;background:#f3faf5}.verde-levato-commessa-row{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px;align-items:end}.verde-levato-new-commessa{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,.55fr) auto;gap:8px;padding-top:8px;border-top:1px solid #c8dbcd}.verde-levato-commessa-help{margin:0;color:#526d5b;font-size:.74rem;font-weight:500}
      body.verde-levato-modal-open{overflow:hidden}
      @media(max-width:760px){.verde-levato-badge{display:none}.verde-levato-hero{grid-template-columns:1fr}.verde-levato-categories{grid-template-columns:1fr}.verde-levato-result{grid-template-columns:1fr}.verde-levato-map{height:46vh;min-height:310px}.verde-levato-form{grid-template-columns:1fr}.verde-levato-form-wide,.verde-levato-location-box,.verde-levato-commessa-box,.verde-levato-form-footer,.verde-levato-form-feedback{grid-column:1}.verde-levato-commessa-row,.verde-levato-new-commessa{grid-template-columns:1fr}.verde-levato-header p{display:none}.verde-levato-header .btn{padding:7px 9px}.verde-levato-result-actions .btn,.verde-levato-result-actions a{flex:1;justify-content:center;text-align:center;text-decoration:none}}
    `;
    document.head.appendChild(style);
  }

  function buildPage() {
    let page = $(PAGE_ID);
    if (page) return page;
    page = document.createElement("section");
    page.id = PAGE_ID;
    page.className = "verde-levato-page hidden";
    page.setAttribute("aria-hidden", "true");
    page.innerHTML = `
      <div class="verde-levato-shell">
        <header class="verde-levato-header">
          <button id="verde-levato-back-btn" class="btn" type="button">← HOME</button>
          <div class="verde-levato-header-copy"><h1>🌱 Verde Levato</h1><p>Cantieri e censimenti inseriti manualmente</p></div>
          <span class="verde-levato-badge">GESTIONE DEDICATA</span>
        </header>
        <section id="verde-levato-hub">
          <div class="verde-levato-hero">
            <article class="verde-levato-card"><h2>Verde Levato</h2><p>La stessa impostazione operativa di Verde Bologna, senza importare cantieri da applicazioni esterne. I dati vengono censiti sul posto e salvati nella sezione dedicata.</p><div class="verde-levato-hero-actions"><button id="verde-levato-new-btn" class="btn btn-primary verde-levato-hidden" type="button">＋ INSERIMENTO DATI</button><button id="verde-levato-export-btn" class="btn verde-levato-hidden" type="button">📊 ESPORTA TUTTI I DATI</button><button id="verde-levato-refresh-btn" class="btn" type="button">↻ AGGIORNA</button></div></article>
            <article class="verde-levato-card verde-levato-access"><h3>Accesso alla gestione</h3><p id="verde-levato-access-summary">Verifica autorizzazioni…</p><div id="verde-levato-admin-list" class="verde-levato-admin-list"></div><button id="verde-levato-admin-btn" class="btn verde-levato-hidden" type="button">👤 AGGIUNGI AMMINISTRATORE</button></article>
          </div>
          <section id="verde-levato-categories" class="verde-levato-categories" aria-label="Categorie Verde Levato"></section>
        </section>
        <section id="verde-levato-browser" class="verde-levato-browser hidden">
          <div class="verde-levato-toolbar"><button id="verde-levato-categories-btn" class="btn" type="button">← CATEGORIE</button><h2 id="verde-levato-category-title">Categoria</h2><button id="verde-levato-category-new-btn" class="btn btn-primary verde-levato-hidden" type="button">＋ AGGIUNGI</button></div>
          <form id="verde-levato-search-form" class="verde-levato-search"><input id="verde-levato-search" type="search" autocomplete="off" placeholder="Cerca nome, codice, via, specie…"><button class="btn" type="submit">CERCA</button></form>
          <p id="verde-levato-status" class="verde-levato-status" role="status">Caricamento dati…</p>
          <section class="verde-levato-map-card"><div class="verde-levato-map-toolbar"><strong>Mappa Verde Levato</strong><button id="verde-levato-map-location-btn" class="btn" type="button">⌖ LA MIA POSIZIONE</button><select id="verde-levato-map-style" aria-label="Vista mappa"><option value="classic">Classica</option><option value="satellite">Satellite</option><option value="hybrid">Ibrida</option></select></div><div id="verde-levato-map" class="verde-levato-map"></div></section>
          <section id="verde-levato-results" class="verde-levato-results"></section>
        </section>
      </div>
      <section id="verde-levato-record-modal" class="verde-levato-modal hidden" aria-hidden="true">
        <div class="verde-levato-modal-card"><header class="verde-levato-modal-head"><h2 id="verde-levato-form-title">Nuovo elemento</h2><button id="verde-levato-form-close" class="btn" type="button">✕</button></header>
          <form id="verde-levato-form" class="verde-levato-form">
            <input name="recordId" type="hidden">
            <label>Tipo *<select name="tipoRecord" required><option value="cantiere">🛠️ Cantiere</option><option value="albero">🌳 Albero censito</option><option value="siepe">🌿 Siepe</option></select></label>
            <label>Codice / numero<input name="codice" maxlength="80" placeholder="Es. LEV-001"></label>
            <section class="verde-levato-commessa-box" data-verde-levato-type="cantiere"><div class="verde-levato-commessa-row"><label>Commessa Verde Levato *<select name="commessaId" required disabled><option value="">Caricamento commesse…</option></select></label><button id="verde-levato-show-new-commessa" class="btn" type="button">＋ NUOVA COMMESSA</button></div><p class="verde-levato-commessa-help">Ogni nuovo cantiere deve essere associato a una commessa. Per adesso le commesse sono gestite soltanto in questo modulo di inserimento dati.</p><div id="verde-levato-new-commessa-panel" class="verde-levato-new-commessa verde-levato-hidden"><label>Nome nuova commessa *<input id="verde-levato-new-commessa-name" maxlength="160" placeholder="Es. Manutenzione Bologna Nord"></label><label>Codice commessa<input id="verde-levato-new-commessa-code" maxlength="80" placeholder="Es. VL-2026-01"></label><button id="verde-levato-save-commessa" class="btn btn-primary" type="button">SALVA COMMESSA</button></div><p id="verde-levato-commessa-feedback" class="verde-levato-form-feedback" role="status"></p></section>
            <label class="verde-levato-form-wide">Denominazione *<input name="denominazione" maxlength="180" required placeholder="Nome del cantiere, albero o siepe"></label>
            <section class="verde-levato-location-box"><div class="verde-levato-location-actions"><button id="verde-levato-use-location" class="btn btn-primary" type="button">📍 LA MIA POSIZIONE</button><strong>Compilazione automatica da GPS</strong></div><p id="verde-levato-location-status" class="verde-levato-location-status">Premi il pulsante sul posto: coordinate e indirizzo verranno compilati automaticamente. Tutti i campi resteranno modificabili.</p></section>
            <label>Latitudine GPS *<input name="gpsY" inputmode="decimal" required placeholder="44.000000"></label><label>Longitudine GPS *<input name="gpsX" inputmode="decimal" required placeholder="11.000000"></label>
            <label>Precisione GPS (metri)<input name="gpsAccuracyM" inputmode="decimal" readonly></label><label>Data/ora rilevazione<input name="gpsDetectedAt" readonly></label>
            <label>Comune<input name="comune" maxlength="120"></label><label>Località / quartiere<input name="localita" maxlength="160"></label>
            <label>Via<input name="via" maxlength="180"></label><label>Civico<input name="civico" maxlength="30"></label>
            <label>CAP<input name="cap" maxlength="20"></label><label>Provincia<input name="provincia" maxlength="120"></label>
            <label>Regione<input name="regione" maxlength="120"></label><label>Paese<input name="paese" maxlength="120"></label>
            <label class="verde-levato-form-wide">Indirizzo completo rilevato<input name="indirizzo" maxlength="320"></label>
            <section class="verde-levato-specific" data-verde-levato-type="cantiere"><label>Superficie (m²)<input name="superficieMq" inputmode="decimal"></label><label>Tipologia intervento<input name="tipologiaIntervento" maxlength="160" placeholder="Sfalcio, potatura, pulizia…"></label><label class="verde-levato-form-wide">Lavorazioni richieste<textarea name="lavorazioniRichieste" maxlength="1500"></textarea></label></section>
            <section class="verde-levato-specific verde-levato-hidden" data-verde-levato-type="albero"><label>Numero albero<input name="numeroAlbero" maxlength="80"></label><label>Specie<input name="specieAlbero" maxlength="160"></label><label>Diametro tronco (cm)<input name="diametroCm" inputmode="decimal"></label><label>Altezza (m)<input name="altezzaAlberoM" inputmode="decimal"></label><label class="verde-levato-form-wide">Stato vegetativo / osservazioni<textarea name="statoVegetativo" maxlength="1000"></textarea></label></section>
            <section class="verde-levato-specific verde-levato-hidden" data-verde-levato-type="siepe"><label>Specie siepe<input name="specieSiepe" maxlength="160"></label><label>Lunghezza (m)<input name="lunghezzaM" inputmode="decimal"></label><label>Altezza (m)<input name="altezzaSiepeM" inputmode="decimal"></label><label>Larghezza (m)<input name="larghezzaM" inputmode="decimal"></label></section>
            <label class="verde-levato-form-wide">Note operative<textarea name="note" maxlength="1800"></textarea></label>
            <p id="verde-levato-form-feedback" class="verde-levato-form-feedback" role="status"></p>
            <footer class="verde-levato-form-footer"><button id="verde-levato-form-cancel" class="btn" type="button">ANNULLA</button><button class="btn btn-primary" type="submit">SALVA</button></footer>
          </form>
        </div>
      </section>
      <section id="verde-levato-admin-modal" class="verde-levato-modal hidden" aria-hidden="true"><div class="verde-levato-modal-card"><header class="verde-levato-modal-head"><h2>Amministratore Verde Levato</h2><button id="verde-levato-admin-close" class="btn" type="button">✕</button></header><form id="verde-levato-admin-form" class="verde-levato-admin-form"><p class="verde-levato-admin-note">Questa email potrà aggiungere e modificare soltanto i dati di Verde Levato. Non diventerà amministratore generale dell’app.</p><label>Email nuovo amministratore<input name="email" type="email" autocomplete="off" required placeholder="operatore@azienda.it"></label><button class="btn btn-primary" type="submit">AGGIUNGI AMMINISTRATORE</button><p id="verde-levato-admin-feedback" class="verde-levato-form-feedback" role="status"></p></form></div></section>`;
    document.body.appendChild(page);
    installEvents();
    return page;
  }

  function setStatus(message, type = "") {
    const node = $("verde-levato-status");
    if (!node) return;
    node.textContent = message;
    node.className = `verde-levato-status ${type}`.trim();
  }

  function recordCategory(record) {
    const value = text(record?.tipoRecord).toLocaleLowerCase("it-IT");
    return CATEGORIES.some((category) => category.id === value) ? value : "cantiere";
  }

  function categoryInfo(categoryId) {
    return CATEGORIES.find((category) => category.id === categoryId) || CATEGORIES[0];
  }

  function renderAccess() {
    const email = normalizeEmail(currentUser()?.email);
    state.globalAdmin = isGlobalAdmin();
    state.canManage = state.globalAdmin || (email && state.adminEmails.includes(email));
    const summary = $("verde-levato-access-summary");
    if (summary) summary.textContent = state.canManage
      ? "Sei autorizzato ad aggiungere e modificare i dati di Verde Levato."
      : "Puoi consultare i dati. L’inserimento è riservato all’amministratore Verde Levato.";
    const list = $("verde-levato-admin-list");
    if (list) list.innerHTML = state.adminEmails.length
      ? state.adminEmails.map((item) => `<span class="verde-levato-admin-chip">${esc(item)}</span>`).join("")
      : '<span class="verde-levato-admin-chip">Nessuna email dedicata configurata</span>';
    [$("verde-levato-new-btn"), $("verde-levato-category-new-btn"), $("verde-levato-export-btn")].forEach((button) => button?.classList.toggle("verde-levato-hidden", !state.canManage));
    $("verde-levato-admin-btn")?.classList.toggle("verde-levato-hidden", !state.globalAdmin);
  }

  async function loadAccess() {
    const firestore = store();
    if (!firestore || !currentUser()) {
      state.adminEmails = [];
      renderAccess();
      return;
    }
    try {
      const snapshot = await firestore.collection(CONFIG_COLLECTION).doc(CONFIG_DOCUMENT).get();
      state.adminEmails = Array.from(new Set((snapshot.exists ? snapshot.data()?.adminEmails : [])
        ?.map(normalizeEmail).filter(Boolean) || [])).sort((a, b) => a.localeCompare(b, "it"));
    } catch (error) {
      console.warn("Verde Levato: impossibile leggere gli amministratori dedicati", error);
      state.adminEmails = [];
    }
    renderAccess();
  }

  function renderCommessaOptions(selectedId = "") {
    const select = $("verde-levato-form")?.elements.commessaId;
    if (!select) return;
    const wanted = text(selectedId || select.value);
    select.innerHTML = `<option value="">${state.commesse.length ? "Seleziona una commessa Verde Levato" : "Nessuna commessa: creane una qui sotto"}</option>${state.commesse.map((commessa) => `<option value="${esc(commessa.id)}">${esc([commessa.codice, commessa.nome].filter(Boolean).join(" · "))}</option>`).join("")}`;
    select.disabled = false;
    if (wanted && state.commesse.some((commessa) => commessa.id === wanted)) select.value = wanted;
  }

  async function loadCommesse(force = false) {
    if (state.commesseLoaded && !force) {
      renderCommessaOptions();
      return state.commesse;
    }
    if (state.commessePromise && !force) return state.commessePromise;
    const firestore = store();
    if (!firestore || !currentUser() || !state.canManage) {
      state.commesse = [];
      state.commesseLoaded = false;
      renderCommessaOptions();
      return state.commesse;
    }
    const select = $("verde-levato-form")?.elements.commessaId;
    if (select) {
      select.disabled = true;
      select.innerHTML = '<option value="">Caricamento commesse…</option>';
    }
    state.commessePromise = (async () => {
      try {
        const snapshot = await firestore.collection(COMMESSE_COLLECTION).get();
        state.commesse = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }))
          .sort((a, b) => text(a.nome).localeCompare(text(b.nome), "it"));
        state.commesseLoaded = true;
        renderCommessaOptions();
        return state.commesse;
      } catch (error) {
        console.error("Verde Levato: caricamento commesse non riuscito", error);
        state.commesseLoaded = false;
        if (select) {
          select.disabled = true;
          select.innerHTML = '<option value="">Impossibile caricare le commesse</option>';
        }
        const feedback = $("verde-levato-commessa-feedback");
        if (feedback) { feedback.textContent = error?.message || "Impossibile caricare le commesse Verde Levato."; feedback.classList.add("error"); }
        throw error;
      } finally {
        state.commessePromise = null;
      }
    })();
    return state.commessePromise;
  }

  function toggleNewCommessaPanel() {
    const panel = $("verde-levato-new-commessa-panel");
    if (!panel) return;
    const willOpen = panel.classList.contains("verde-levato-hidden");
    panel.classList.toggle("verde-levato-hidden", !willOpen);
    if (willOpen) $("verde-levato-new-commessa-name")?.focus();
  }

  async function saveCommessa() {
    const feedback = $("verde-levato-commessa-feedback");
    const nome = text($("verde-levato-new-commessa-name")?.value);
    const codice = text($("verde-levato-new-commessa-code")?.value);
    if (!state.canManage || !currentUser() || !store()) {
      if (feedback) { feedback.textContent = "Autorizzazione Verde Levato non disponibile."; feedback.classList.add("error"); }
      return;
    }
    if (!nome) {
      if (feedback) { feedback.textContent = "Inserisci il nome della nuova commessa."; feedback.classList.add("error"); }
      return;
    }
    const duplicate = state.commesse.find((commessa) => text(commessa.nome).toLocaleLowerCase("it-IT") === nome.toLocaleLowerCase("it-IT") || (codice && text(commessa.codice).toLocaleLowerCase("it-IT") === codice.toLocaleLowerCase("it-IT")));
    if (duplicate) {
      renderCommessaOptions(duplicate.id);
      if (feedback) { feedback.textContent = "La commessa esiste già ed è stata selezionata."; feedback.classList.remove("error"); }
      return;
    }
    const button = $("verde-levato-save-commessa");
    if (button) button.disabled = true;
    if (feedback) { feedback.textContent = "Salvataggio nuova commessa…"; feedback.classList.remove("error"); }
    try {
      const reference = store().collection(COMMESSE_COLLECTION).doc();
      const user = currentUser();
      const payload = {
        nome, codice, source: "MANUALE_VERDE_LEVATO", createdAt: firestoreTimestamp(), updatedAt: firestoreTimestamp(),
        createdByUid: user?.uid || "", createdByEmail: normalizeEmail(user?.email), updatedByUid: user?.uid || "", updatedByEmail: normalizeEmail(user?.email)
      };
      await reference.set(payload, { merge: true });
      state.commesse.push({ id: reference.id, ...payload });
      state.commesse.sort((a, b) => text(a.nome).localeCompare(text(b.nome), "it"));
      state.commesseLoaded = true;
      renderCommessaOptions(reference.id);
      if ($("verde-levato-new-commessa-name")) $("verde-levato-new-commessa-name").value = "";
      if ($("verde-levato-new-commessa-code")) $("verde-levato-new-commessa-code").value = "";
      $("verde-levato-new-commessa-panel")?.classList.add("verde-levato-hidden");
      if (feedback) feedback.textContent = `Commessa “${nome}” creata e associata al nuovo cantiere.`;
    } catch (error) {
      console.error("Verde Levato: salvataggio commessa non riuscito", error);
      if (feedback) { feedback.textContent = error?.message || "Impossibile salvare la nuova commessa."; feedback.classList.add("error"); }
    } finally {
      if (button) button.disabled = false;
    }
  }

  async function loadRecords() {
    if (state.loading) return;
    const firestore = store();
    if (!firestore || !currentUser()) {
      state.records = [];
      setStatus("Accedi per consultare Verde Levato.", "warning");
      renderAll();
      return;
    }
    state.loading = true;
    setStatus("Caricamento censimenti Verde Levato…");
    try {
      const snapshot = await firestore.collection(RECORDS_COLLECTION).get();
      state.records = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() })).sort((a, b) => {
        const aTime = a.updatedAt?.toMillis?.() || a.updatedAt?.getTime?.() || 0;
        const bTime = b.updatedAt?.toMillis?.() || b.updatedAt?.getTime?.() || 0;
        return bTime - aTime || text(a.denominazione).localeCompare(text(b.denominazione), "it");
      });
      setStatus(`${state.records.length} elementi caricati. Nessun dato proviene da app esterne.`);
    } catch (error) {
      console.error("Verde Levato: caricamento non riuscito", error);
      setStatus(error?.message || "Impossibile caricare i dati Verde Levato.", "error");
    } finally {
      state.loading = false;
      renderAll();
    }
  }

  function renderCategories() {
    const host = $("verde-levato-categories");
    if (!host) return;
    host.innerHTML = CATEGORIES.map((category) => {
      const count = state.records.filter((record) => recordCategory(record) === category.id).length;
      return `<button class="verde-levato-category" type="button" data-verde-levato-category="${category.id}"><span class="verde-levato-category-icon">${category.icon}</span><h3>${esc(category.title)}</h3><p>${esc(category.description)}</p><strong>${count} elementi</strong><span class="btn btn-primary">APRI</span></button>`;
    }).join("");
    host.querySelectorAll("[data-verde-levato-category]").forEach((button) => button.addEventListener("click", () => openCategory(button.dataset.verdeLevatoCategory)));
  }

  function matchingRecords() {
    const needle = state.query.toLocaleLowerCase("it-IT");
    return state.records.filter((record) => recordCategory(record) === state.category).filter((record) => {
      if (!needle) return true;
      return [record.denominazione, record.codice, record.comune, record.localita, record.via, record.indirizzo, record.specieAlbero, record.specieSiepe, record.note]
        .some((value) => text(value).toLocaleLowerCase("it-IT").includes(needle));
    });
  }

  function detailPairs(record) {
    const common = [
      ["Codice", record.codice], ["Comune", record.comune], ["Indirizzo", record.indirizzo || [record.via, record.civico].filter(Boolean).join(" ")],
      ["Coordinate", Number.isFinite(Number(record.gpsY)) && Number.isFinite(Number(record.gpsX)) ? `${Number(record.gpsY).toFixed(6)}, ${Number(record.gpsX).toFixed(6)}` : ""], ["Note", record.note]
    ];
    if (recordCategory(record) === "albero") common.splice(1, 0, ["Numero albero", record.numeroAlbero], ["Specie", record.specieAlbero], ["Diametro", record.diametroCm ? `${record.diametroCm} cm` : ""], ["Altezza", record.altezzaAlberoM ? `${record.altezzaAlberoM} m` : ""]);
    if (recordCategory(record) === "siepe") common.splice(1, 0, ["Specie", record.specieSiepe], ["Lunghezza", record.lunghezzaM ? `${record.lunghezzaM} m` : ""], ["Altezza", record.altezzaSiepeM ? `${record.altezzaSiepeM} m` : ""], ["Larghezza", record.larghezzaM ? `${record.larghezzaM} m` : ""]);
    if (recordCategory(record) === "cantiere") common.splice(1, 0, ["Superficie", record.superficieMq ? `${record.superficieMq} m²` : ""], ["Intervento", record.tipologiaIntervento], ["Lavorazioni", record.lavorazioniRichieste]);
    return common.filter(([, value]) => text(value));
  }

  function renderResults() {
    const host = $("verde-levato-results");
    if (!host || !state.category) return;
    const category = categoryInfo(state.category);
    const records = matchingRecords();
    if (!records.length) {
      host.innerHTML = `<p class="verde-levato-empty">Nessun elemento in “${esc(category.title)}”${state.query ? " per questa ricerca" : ""}.</p>`;
      renderMap(records);
      return;
    }
    host.innerHTML = records.map((record) => {
      const lat = Number(record.gpsY), lon = Number(record.gpsX);
      const navigable = Number.isFinite(lat) && Number.isFinite(lon);
      const pairs = detailPairs(record).slice(0, 8);
      return `<article class="verde-levato-result" data-verde-levato-record="${esc(record.id)}"><div><h3>${category.icon} ${esc(record.denominazione || category.title)}</h3><p>${esc([record.codice, record.comune, record.via].filter(Boolean).join(" · "))}</p></div><div class="verde-levato-result-actions">${navigable ? `<a class="btn" href="https://www.google.com/maps/dir/?api=1&destination=${lat},${lon}" target="_blank" rel="noopener">NAVIGA</a><button class="btn" type="button" data-verde-levato-map="${esc(record.id)}">MAPPA</button>` : ""}${state.canManage ? `<button class="btn" type="button" data-verde-levato-edit="${esc(record.id)}">MODIFICA</button>` : ""}</div><div class="verde-levato-result-details">${pairs.map(([label, value]) => `<div><span>${esc(label)}</span><strong>${esc(value)}</strong></div>`).join("")}</div></article>`;
    }).join("");
    host.querySelectorAll("[data-verde-levato-edit]").forEach((button) => button.addEventListener("click", () => openRecordForm(state.records.find((record) => record.id === button.dataset.verdeLevatoEdit))));
    host.querySelectorAll("[data-verde-levato-map]").forEach((button) => button.addEventListener("click", () => focusRecord(button.dataset.verdeLevatoMap)));
    renderMap(records);
  }

  function renderAll() {
    renderCategories();
    if (state.category) renderResults();
  }

  function ensureMap() {
    if (state.map) return state.map;
    const node = $("verde-levato-map");
    if (!node || !window.L) return null;
    state.map = L.map(node, { zoomControl: true, zoomAnimation: false, fadeAnimation: false, markerZoomAnimation: false }).setView(DEFAULT_CENTER, 14);
    state.markersLayer = L.layerGroup().addTo(state.map);
    applyMapStyle($("verde-levato-map-style")?.value || "classic");
    return state.map;
  }

  function applyMapStyle(style) {
    if (!state.map) return;
    state.baseLayer?.remove?.();
    state.labelsLayer?.remove?.();
    state.labelsLayer = null;
    const selected = style === "satellite" || style === "hybrid" ? TILE_LAYERS.satellite : TILE_LAYERS.classic;
    state.baseLayer = L.tileLayer(selected.url, selected.options).addTo(state.map);
    if (style === "hybrid") state.labelsLayer = L.tileLayer(TILE_LAYERS.labels.url, TILE_LAYERS.labels.options).addTo(state.map);
  }

  function renderMap(records = matchingRecords()) {
    const map = ensureMap();
    if (!map || !state.markersLayer) return;
    state.markersLayer.clearLayers();
    const bounds = L.latLngBounds([]);
    records.forEach((record) => {
      const lat = Number(record.gpsY), lon = Number(record.gpsX);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) return;
      const category = categoryInfo(recordCategory(record));
      const icon = L.divIcon({ className: "verde-levato-marker-wrap", html: `<span class="verde-levato-marker is-${category.id}">${category.icon}</span>`, iconSize: [34, 34], iconAnchor: [17, 17], popupAnchor: [0, -17] });
      L.marker([lat, lon], { icon }).addTo(state.markersLayer).bindPopup(`<strong>${esc(record.denominazione || category.title)}</strong><br>${esc(record.indirizzo || record.comune || "")}`);
      bounds.extend([lat, lon]);
    });
    if (bounds.isValid()) map.fitBounds(bounds, { padding: [34, 34], maxZoom: 17, animate: false });
    window.setTimeout(() => map.invalidateSize(false), 0);
  }

  function focusRecord(recordId) {
    const record = state.records.find((item) => item.id === recordId);
    const map = ensureMap();
    const lat = Number(record?.gpsY), lon = Number(record?.gpsX);
    if (!map || !Number.isFinite(lat) || !Number.isFinite(lon)) return;
    map.setView([lat, lon], 18, { animate: false });
    $("verde-levato-map")?.scrollIntoView?.({ behavior: "smooth", block: "center" });
  }

  function openCategory(categoryId) {
    if (!CATEGORIES.some((category) => category.id === categoryId)) return;
    state.category = categoryId;
    state.query = "";
    if ($("verde-levato-search")) $("verde-levato-search").value = "";
    $("verde-levato-hub")?.classList.add("hidden");
    $("verde-levato-browser")?.classList.remove("hidden");
    const category = categoryInfo(categoryId);
    if ($("verde-levato-category-title")) $("verde-levato-category-title").textContent = `${category.icon} ${category.title}`;
    renderResults();
  }

  function showCategories() {
    state.category = "";
    state.query = "";
    $("verde-levato-browser")?.classList.add("hidden");
    $("verde-levato-hub")?.classList.remove("hidden");
    renderCategories();
  }

  function showMapUserLocation() {
    if (!navigator.geolocation) { setStatus("Geolocalizzazione non supportata.", "error"); return; }
    const button = $("verde-levato-map-location-btn");
    if (button) button.disabled = true;
    navigator.geolocation.getCurrentPosition((position) => {
      const map = ensureMap();
      const lat = Number(position.coords.latitude), lon = Number(position.coords.longitude), accuracy = Number(position.coords.accuracy);
      if (!map || !Number.isFinite(lat) || !Number.isFinite(lon)) return;
      const point = L.latLng(lat, lon);
      if (!state.locationMarker) state.locationMarker = L.circleMarker(point, { radius: 9, color: "#fff", weight: 3, fillColor: "#1268e8", fillOpacity: 1 }).addTo(map).bindPopup("<strong>La mia posizione</strong>"); else state.locationMarker.setLatLng(point).addTo(map);
      if (Number.isFinite(accuracy) && accuracy > 0) {
        if (!state.locationAccuracy) state.locationAccuracy = L.circle(point, { radius: accuracy, color: "#1268e8", weight: 1, fillOpacity: .08 }).addTo(map); else state.locationAccuracy.setLatLng(point).setRadius(accuracy).addTo(map);
      }
      map.setView(point, Math.max(map.getZoom(), 17), { animate: false });
      state.locationMarker.openPopup();
      if (button) button.disabled = false;
    }, (error) => { if (button) button.disabled = false; setStatus(error?.code === 1 ? "Permesso posizione negato." : "Impossibile rilevare la posizione.", "error"); }, { enableHighAccuracy: true, timeout: 15000, maximumAge: 15000 });
  }

  function setLocationStatus(message, error = false) {
    const node = $("verde-levato-location-status");
    if (!node) return;
    node.textContent = message;
    node.style.color = error ? "#9b281f" : "#4d6c56";
  }

  function fillAddressFields(form, result) {
    const address = result?.address || {};
    const values = {
      comune: address.city || address.town || address.village || address.municipality || "",
      localita: address.suburb || address.neighbourhood || address.hamlet || address.city_district || "",
      via: address.road || address.pedestrian || address.path || address.cycleway || "",
      civico: address.house_number || "",
      cap: address.postcode || "",
      provincia: address.county || address.province || "",
      regione: address.state || address.region || "",
      paese: address.country || "",
      indirizzo: result?.display_name || ""
    };
    Object.entries(values).forEach(([name, value]) => { if (form.elements[name] && value) form.elements[name].value = value; });
    if (!text(form.elements.denominazione?.value)) {
      const detectedName = address.amenity || address.leisure || address.park || address.building || address.road || values.localita || values.comune;
      if (detectedName) form.elements.denominazione.value = detectedName;
    }
  }

  async function reverseGeocode(form, lat, lon) {
    if (navigator.onLine === false) {
      setLocationStatus("Coordinate GPS compilate. Senza connessione inserisci manualmente Comune e Via.");
      return;
    }
    setLocationStatus("Posizione rilevata. Cerco automaticamente indirizzo e dati territoriali…");
    try {
      const url = `https://nominatim.openstreetmap.org/reverse?format=jsonv2&lat=${encodeURIComponent(lat)}&lon=${encodeURIComponent(lon)}&zoom=18&addressdetails=1&accept-language=it`;
      const response = await fetch(url, { headers: { Accept: "application/json" } });
      if (!response.ok) throw new Error(`Servizio indirizzi non disponibile (${response.status})`);
      const result = await response.json();
      fillAddressFields(form, result);
      setLocationStatus("Coordinate, Comune, Via e altri dati territoriali compilati. Controllali e completa i dati tecnici manuali.");
    } catch (error) {
      setLocationStatus(`${error?.message || "Indirizzo non rilevato"}. Le coordinate sono state mantenute; compila l’indirizzo manualmente.`, true);
    }
  }

  function useCurrentLocation() {
    const form = $("verde-levato-form");
    if (!form || !navigator.geolocation) { setLocationStatus("Geolocalizzazione non supportata da questo dispositivo.", true); return; }
    const button = $("verde-levato-use-location");
    if (button) button.disabled = true;
    setLocationStatus("Rilevazione GPS ad alta precisione…");
    navigator.geolocation.getCurrentPosition(async (position) => {
      const lat = Number(position.coords.latitude), lon = Number(position.coords.longitude), accuracy = Number(position.coords.accuracy);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) { setLocationStatus("Coordinate ricevute non valide.", true); if (button) button.disabled = false; return; }
      form.elements.gpsY.value = lat.toFixed(7);
      form.elements.gpsX.value = lon.toFixed(7);
      form.elements.gpsAccuracyM.value = Number.isFinite(accuracy) ? accuracy.toFixed(1) : "";
      form.elements.gpsDetectedAt.value = new Date(position.timestamp || Date.now()).toLocaleString("it-IT");
      form.dataset.coordinateSource = "DEVICE_GPS";
      await reverseGeocode(form, lat, lon);
      if (button) button.disabled = false;
    }, (error) => { if (button) button.disabled = false; setLocationStatus(error?.code === 1 ? "Permesso posizione negato. Abilita la posizione nelle impostazioni del telefono." : "Impossibile rilevare la posizione GPS. Riprova all’aperto.", true); }, { enableHighAccuracy: true, timeout: 18000, maximumAge: 0 });
  }

  function syncTypeFields() {
    const form = $("verde-levato-form");
    const type = form?.elements.tipoRecord?.value || "cantiere";
    form?.querySelectorAll("[data-verde-levato-type]").forEach((section) => section.classList.toggle("verde-levato-hidden", section.dataset.verdeLevatoType !== type));
    if (form?.elements.commessaId) form.elements.commessaId.required = type === "cantiere";
    if (type === "cantiere" && state.canManage && !state.commesseLoaded) loadCommesse().catch(() => {});
  }

  async function openRecordForm(record = null) {
    if (!state.canManage) { window.alert("Solo l’amministratore Verde Levato può aggiungere o modificare elementi."); return; }
    const form = $("verde-levato-form");
    if (!form) return;
    form.reset();
    form.dataset.coordinateSource = record?.coordinateSource || "MANUAL";
    const fields = ["recordId", "tipoRecord", "codice", "commessaId", "denominazione", "gpsY", "gpsX", "gpsAccuracyM", "gpsDetectedAt", "comune", "localita", "via", "civico", "cap", "provincia", "regione", "paese", "indirizzo", "superficieMq", "tipologiaIntervento", "lavorazioniRichieste", "numeroAlbero", "specieAlbero", "diametroCm", "altezzaAlberoM", "statoVegetativo", "specieSiepe", "lunghezzaM", "altezzaSiepeM", "larghezzaM", "note"];
    fields.forEach((name) => {
      if (!form.elements[name] || (!record && name !== "recordId")) return;
      form.elements[name].value = name === "recordId" ? (record?.id || "") : (record?.[name] ?? "");
    });
    if (!record) form.elements.tipoRecord.value = state.category || "cantiere";
    if ($("verde-levato-form-title")) $("verde-levato-form-title").textContent = record ? "Modifica elemento Verde Levato" : "Nuovo elemento Verde Levato";
    if ($("verde-levato-form-feedback")) $("verde-levato-form-feedback").textContent = "";
    if ($("verde-levato-commessa-feedback")) { $("verde-levato-commessa-feedback").textContent = ""; $("verde-levato-commessa-feedback").classList.remove("error"); }
    $("verde-levato-new-commessa-panel")?.classList.add("verde-levato-hidden");
    setLocationStatus("Premi LA MIA POSIZIONE sul posto per compilare automaticamente coordinate e indirizzo.");
    syncTypeFields();
    $("verde-levato-record-modal")?.classList.remove("hidden");
    $("verde-levato-record-modal")?.setAttribute("aria-hidden", "false");
    document.body.classList.add("verde-levato-modal-open");
    if (form.elements.tipoRecord.value === "cantiere") {
      await loadCommesse().catch(() => []);
      renderCommessaOptions(record?.commessaId || "");
    }
  }

  function closeRecordForm() {
    $("verde-levato-record-modal")?.classList.add("hidden");
    $("verde-levato-record-modal")?.setAttribute("aria-hidden", "true");
    document.body.classList.remove("verde-levato-modal-open");
  }

  function optionalNumber(form, name) {
    const raw = text(form.elements[name]?.value);
    return raw ? numberValue(raw) : null;
  }

  function buildRecordPayload(form) {
    const user = currentUser();
    const tipoRecord = text(form.elements.tipoRecord.value);
    const selectedCommessa = tipoRecord === "cantiere"
      ? state.commesse.find((commessa) => commessa.id === text(form.elements.commessaId.value))
      : null;
    const payload = {
      tipoRecord,
      codice: text(form.elements.codice.value),
      commessaId: selectedCommessa?.id || "",
      commessaNome: text(selectedCommessa?.nome),
      commessaCodice: text(selectedCommessa?.codice),
      denominazione: text(form.elements.denominazione.value),
      gpsY: numberValue(form.elements.gpsY.value),
      gpsX: numberValue(form.elements.gpsX.value),
      gpsAccuracyM: optionalNumber(form, "gpsAccuracyM"),
      gpsDetectedAt: text(form.elements.gpsDetectedAt.value),
      coordinateSource: form.dataset.coordinateSource || "MANUAL",
      comune: text(form.elements.comune.value), localita: text(form.elements.localita.value), via: text(form.elements.via.value), civico: text(form.elements.civico.value), cap: text(form.elements.cap.value), provincia: text(form.elements.provincia.value), regione: text(form.elements.regione.value), paese: text(form.elements.paese.value), indirizzo: text(form.elements.indirizzo.value),
      superficieMq: optionalNumber(form, "superficieMq"), tipologiaIntervento: text(form.elements.tipologiaIntervento.value), lavorazioniRichieste: text(form.elements.lavorazioniRichieste.value),
      numeroAlbero: text(form.elements.numeroAlbero.value), specieAlbero: text(form.elements.specieAlbero.value), diametroCm: optionalNumber(form, "diametroCm"), altezzaAlberoM: optionalNumber(form, "altezzaAlberoM"), statoVegetativo: text(form.elements.statoVegetativo.value),
      specieSiepe: text(form.elements.specieSiepe.value), lunghezzaM: optionalNumber(form, "lunghezzaM"), altezzaSiepeM: optionalNumber(form, "altezzaSiepeM"), larghezzaM: optionalNumber(form, "larghezzaM"),
      note: text(form.elements.note.value), source: "MANUALE_VERDE_LEVATO", updatedAt: firestoreTimestamp(), updatedByUid: user?.uid || "", updatedByEmail: normalizeEmail(user?.email), updatedByName: text(user?.displayName || user?.email || "Operatore")
    };
    return payload;
  }

  async function saveRecord(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const feedback = $("verde-levato-form-feedback");
    if (!state.canManage || !currentUser() || !store()) { if (feedback) { feedback.textContent = "Autorizzazione Verde Levato non disponibile."; feedback.classList.add("error"); } return; }
    const payload = buildRecordPayload(form);
    if (!payload.denominazione || !CATEGORIES.some((category) => category.id === payload.tipoRecord)) { if (feedback) { feedback.textContent = "Inserisci tipo e denominazione."; feedback.classList.add("error"); } return; }
    if (payload.tipoRecord === "cantiere" && !payload.commessaId) { if (feedback) { feedback.textContent = "Seleziona o crea la commessa Verde Levato da associare al cantiere."; feedback.classList.add("error"); } return; }
    if (!Number.isFinite(payload.gpsY) || !Number.isFinite(payload.gpsX) || Math.abs(payload.gpsY) > 90 || Math.abs(payload.gpsX) > 180) { if (feedback) { feedback.textContent = "Rileva o inserisci coordinate GPS valide."; feedback.classList.add("error"); } return; }
    const submit = form.querySelector('button[type="submit"]');
    if (submit) submit.disabled = true;
    if (feedback) { feedback.textContent = "Salvataggio Verde Levato…"; feedback.classList.remove("error"); }
    try {
      const recordId = text(form.elements.recordId.value);
      const reference = recordId ? store().collection(RECORDS_COLLECTION).doc(recordId) : store().collection(RECORDS_COLLECTION).doc();
      if (!recordId) Object.assign(payload, { createdAt: firestoreTimestamp(), createdByUid: currentUser()?.uid || "", createdByEmail: normalizeEmail(currentUser()?.email) });
      await reference.set(payload, { merge: true });
      if (feedback) feedback.textContent = navigator.onLine === false ? "Salvato sul dispositivo: sarà sincronizzato appena torni online." : "Elemento Verde Levato salvato.";
      await loadRecords();
      closeRecordForm();
      if (state.category !== payload.tipoRecord) openCategory(payload.tipoRecord); else renderResults();
    } catch (error) {
      console.error("Verde Levato: salvataggio non riuscito", error);
      if (feedback) { feedback.textContent = error?.message || "Salvataggio non riuscito."; feedback.classList.add("error"); }
    } finally { if (submit) submit.disabled = false; }
  }

  function exportDate(value) {
    const date = value?.toDate?.() || (value instanceof Date ? value : null);
    return date && Number.isFinite(date.getTime()) ? date.toLocaleString("it-IT") : text(value);
  }

  function exportRecordRow(record) {
    return {
      "ID record": record.id || "",
      "Commessa": record.commessaNome || "",
      "Codice commessa": record.commessaCodice || "",
      "ID commessa": record.commessaId || "",
      "Tipo elemento": categoryInfo(recordCategory(record)).title,
      "Codice / numero": record.codice || "",
      "Denominazione": record.denominazione || "",
      "Latitudine GPS": record.gpsY ?? "",
      "Longitudine GPS": record.gpsX ?? "",
      "Precisione GPS (m)": record.gpsAccuracyM ?? "",
      "Data/ora rilevazione GPS": record.gpsDetectedAt || "",
      "Origine coordinate": record.coordinateSource || "",
      "Comune": record.comune || "",
      "Località / quartiere": record.localita || "",
      "Via": record.via || "",
      "Civico": record.civico || "",
      "CAP": record.cap || "",
      "Provincia": record.provincia || "",
      "Regione": record.regione || "",
      "Paese": record.paese || "",
      "Indirizzo completo": record.indirizzo || "",
      "Superficie (m²)": record.superficieMq ?? "",
      "Tipologia intervento": record.tipologiaIntervento || "",
      "Lavorazioni richieste": record.lavorazioniRichieste || "",
      "Numero albero": record.numeroAlbero || "",
      "Specie albero": record.specieAlbero || "",
      "Diametro tronco (cm)": record.diametroCm ?? "",
      "Altezza albero (m)": record.altezzaAlberoM ?? "",
      "Stato vegetativo / osservazioni": record.statoVegetativo || "",
      "Specie siepe": record.specieSiepe || "",
      "Lunghezza siepe (m)": record.lunghezzaM ?? "",
      "Altezza siepe (m)": record.altezzaSiepeM ?? "",
      "Larghezza siepe (m)": record.larghezzaM ?? "",
      "Note operative": record.note || "",
      "Origine dato": record.source || "",
      "Creato il": exportDate(record.createdAt),
      "Creato da": record.createdByEmail || record.createdByUid || "",
      "Aggiornato il": exportDate(record.updatedAt),
      "Aggiornato da": record.updatedByName || record.updatedByEmail || record.updatedByUid || ""
    };
  }

  async function exportAllData() {
    if (!state.canManage || !currentUser()) {
      window.alert("L’esportazione è riservata all’amministratore Verde Levato.");
      return;
    }
    const button = $("verde-levato-export-btn");
    if (button) button.disabled = true;
    setStatus("Preparazione del foglio Excel con tutti i dati Verde Levato…");
    try {
      await loadCommesse();
      if (!window.XLSX) await window.HeraHeavyLibs?.ensure?.("xlsx");
      if (!window.XLSX?.utils?.json_to_sheet || typeof window.XLSX.writeFile !== "function") throw new Error("Libreria Excel non disponibile");
      const dataRows = state.records.map(exportRecordRow);
      const emptyDataRow = exportRecordRow({});
      const dataSheet = window.XLSX.utils.json_to_sheet(dataRows.length ? dataRows : [], { header: Object.keys(emptyDataRow) });
      dataSheet["!cols"] = Object.keys(emptyDataRow).map((header) => ({ wch: Math.min(42, Math.max(14, header.length + 2)) }));
      if (dataSheet["!ref"]) dataSheet["!autofilter"] = { ref: dataSheet["!ref"] };

      const commesseRows = state.commesse.map((commessa) => ({
        "ID commessa": commessa.id,
        "Codice commessa": commessa.codice || "",
        "Nome commessa": commessa.nome || "",
        "Cantieri associati": state.records.filter((record) => recordCategory(record) === "cantiere" && record.commessaId === commessa.id).length,
        "Creata il": exportDate(commessa.createdAt),
        "Creata da": commessa.createdByEmail || commessa.createdByUid || "",
        "Aggiornata il": exportDate(commessa.updatedAt),
        "Aggiornata da": commessa.updatedByEmail || commessa.updatedByUid || ""
      }));
      const commesseHeaders = ["ID commessa", "Codice commessa", "Nome commessa", "Cantieri associati", "Creata il", "Creata da", "Aggiornata il", "Aggiornata da"];
      const commesseSheet = window.XLSX.utils.json_to_sheet(commesseRows, { header: commesseHeaders });
      commesseSheet["!cols"] = commesseHeaders.map((header) => ({ wch: Math.min(36, Math.max(16, header.length + 2)) }));
      if (commesseSheet["!ref"]) commesseSheet["!autofilter"] = { ref: commesseSheet["!ref"] };

      const workbook = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(workbook, dataSheet, "Dati completi");
      window.XLSX.utils.book_append_sheet(workbook, commesseSheet, "Commesse");
      const timestamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
      window.XLSX.writeFile(workbook, `verde_levato_dati_completi_${timestamp}.xlsx`);
      setStatus(`${state.records.length} elementi e ${state.commesse.length} commesse esportati in Excel.`);
    } catch (error) {
      console.error("Verde Levato: esportazione Excel non riuscita", error);
      setStatus(error?.message || "Esportazione Excel non riuscita.", "error");
      window.alert("Impossibile esportare i dati Verde Levato. Controlla la connessione e riprova.");
    } finally {
      if (button) button.disabled = false;
    }
  }

  function openAdminModal() {
    if (!state.globalAdmin) return;
    const form = $("verde-levato-admin-form");
    form?.reset();
    if ($("verde-levato-admin-feedback")) $("verde-levato-admin-feedback").textContent = "";
    $("verde-levato-admin-modal")?.classList.remove("hidden");
    $("verde-levato-admin-modal")?.setAttribute("aria-hidden", "false");
    document.body.classList.add("verde-levato-modal-open");
  }

  function closeAdminModal() {
    $("verde-levato-admin-modal")?.classList.add("hidden");
    $("verde-levato-admin-modal")?.setAttribute("aria-hidden", "true");
    document.body.classList.remove("verde-levato-modal-open");
  }

  async function addDedicatedAdmin(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const feedback = $("verde-levato-admin-feedback");
    const email = normalizeEmail(form.elements.email.value);
    if (!state.globalAdmin || !store() || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { if (feedback) { feedback.textContent = "Inserisci un’email valida. L’operazione è riservata all’amministratore generale."; feedback.classList.add("error"); } return; }
    const nextEmails = Array.from(new Set([...state.adminEmails, email])).sort((a, b) => a.localeCompare(b, "it"));
    try {
      await store().collection(CONFIG_COLLECTION).doc(CONFIG_DOCUMENT).set({ adminEmails: nextEmails, updatedAt: firestoreTimestamp(), updatedByUid: currentUser()?.uid || "", updatedByEmail: normalizeEmail(currentUser()?.email) }, { merge: true });
      state.adminEmails = nextEmails;
      renderAccess();
      if (feedback) { feedback.textContent = `${email} ora è amministratore soltanto di Verde Levato.`; feedback.classList.remove("error"); }
      form.elements.email.value = "";
    } catch (error) {
      if (feedback) { feedback.textContent = error?.message || "Impossibile aggiungere l’amministratore."; feedback.classList.add("error"); }
    }
  }

  async function openPage() {
    $("menu-close-btn")?.click();
    $("home-page")?.classList.add("hidden");
    const page = buildPage();
    page.classList.remove("hidden");
    page.setAttribute("aria-hidden", "false");
    showCategories();
    await Promise.all([loadAccess(), loadRecords()]);
  }

  function closePage() {
    closeRecordForm();
    closeAdminModal();
    const page = $(PAGE_ID);
    page?.classList.add("hidden");
    page?.setAttribute("aria-hidden", "true");
    $("home-page")?.classList.remove("hidden");
  }

  function addMenuButton() {
    if ($(MENU_BUTTON_ID)) return;
    const verdeBologna = $("open-verde-bologna-btn");
    const anchor = verdeBologna || $("open-green-areas-btn") || document.querySelector("#side-menu .menu-section");
    if (!anchor) return;
    const button = document.createElement("button");
    button.id = MENU_BUTTON_ID;
    button.className = "btn menu-title-btn";
    button.type = "button";
    button.innerHTML = '<span class="menu-item-icon" aria-hidden="true">🌱</span>Verde Levato';
    if (verdeBologna) verdeBologna.insertAdjacentElement("afterend", button); else if (anchor.matches?.("button")) anchor.insertAdjacentElement("beforebegin", button); else anchor.appendChild(button);
    button.addEventListener("click", openPage);
  }

  function installEvents() {
    $("verde-levato-back-btn")?.addEventListener("click", closePage);
    $("verde-levato-refresh-btn")?.addEventListener("click", () => Promise.all([loadAccess(), loadRecords()]));
    $("verde-levato-new-btn")?.addEventListener("click", () => openRecordForm());
    $("verde-levato-export-btn")?.addEventListener("click", exportAllData);
    $("verde-levato-category-new-btn")?.addEventListener("click", () => openRecordForm());
    $("verde-levato-categories-btn")?.addEventListener("click", showCategories);
    $("verde-levato-search-form")?.addEventListener("submit", (event) => { event.preventDefault(); state.query = text($("verde-levato-search")?.value); renderResults(); });
    $("verde-levato-search")?.addEventListener("input", (event) => { if (!event.currentTarget.value) { state.query = ""; renderResults(); } });
    $("verde-levato-map-location-btn")?.addEventListener("click", showMapUserLocation);
    $("verde-levato-map-style")?.addEventListener("change", (event) => applyMapStyle(event.currentTarget.value));
    $("verde-levato-use-location")?.addEventListener("click", useCurrentLocation);
    $("verde-levato-form")?.elements.tipoRecord?.addEventListener("change", syncTypeFields);
    $("verde-levato-show-new-commessa")?.addEventListener("click", toggleNewCommessaPanel);
    $("verde-levato-save-commessa")?.addEventListener("click", saveCommessa);
    $("verde-levato-form")?.addEventListener("submit", saveRecord);
    $("verde-levato-form-close")?.addEventListener("click", closeRecordForm);
    $("verde-levato-form-cancel")?.addEventListener("click", closeRecordForm);
    $("verde-levato-admin-btn")?.addEventListener("click", openAdminModal);
    $("verde-levato-admin-close")?.addEventListener("click", closeAdminModal);
    $("verde-levato-admin-form")?.addEventListener("submit", addDedicatedAdmin);
    window.addEventListener("resize", () => state.map?.invalidateSize?.(false), { passive: true });
    document.addEventListener("keydown", (event) => { if (event.key !== "Escape" || $(PAGE_ID)?.classList.contains("hidden")) return; if (!$("verde-levato-record-modal")?.classList.contains("hidden")) closeRecordForm(); else if (!$("verde-levato-admin-modal")?.classList.contains("hidden")) closeAdminModal(); else if (state.category) showCategories(); else closePage(); });
  }

  function install() {
    injectStyle();
    buildPage();
    renderCategories();
    addMenuButton();
    window.HeraVerdeLevato = Object.freeze({ open: openPage, close: closePage, useCurrentLocation, exportAllData, categories: CATEGORIES.map(({ id, title }) => ({ id, title })), collections: Object.freeze({ records: RECORDS_COLLECTION, commesse: COMMESSE_COLLECTION, config: CONFIG_COLLECTION }) });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true }); else install();
})();

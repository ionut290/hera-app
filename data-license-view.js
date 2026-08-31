(() => {
  "use strict";

  const VIEW_ID = "data-license-page";
  const STYLE_ID = "data-license-view-style";
  const VERIFIED_ON = "31 agosto 2026";

  const SHARED_MAP_SOURCES = Object.freeze([
    {
      name: "OpenStreetMap",
      role: "Dati cartografici e mappa classica",
      license: "Open Data Commons Open Database License 1.0 (ODbL)",
      sourceUrl: "https://www.openstreetmap.org/copyright/it",
      licenseUrl: "https://opendatacommons.org/licenses/odbl/1-0/",
      attribution: "© contributori OpenStreetMap",
      conditions: [
        "È consentito copiare, usare e adattare i dati, anche per finalità commerciali.",
        "Occorre attribuire OpenStreetMap e i suoi contributori e indicare la licenza ODbL.",
        "Le banche dati derivate distribuite pubblicamente possono essere soggette all’obbligo di condivisione alle stesse condizioni."
      ]
    },
    {
      name: "Esri World Imagery",
      role: "Sfondo satellitare e relative etichette",
      license: "Servizio cartografico soggetto ai termini Esri",
      sourceUrl: "https://www.esri.com/en-us/legal/terms/full-master-agreement",
      attribution: "Tiles © Esri e fornitori indicati nella mappa",
      conditions: [
        "Le immagini satellitari sono utilizzate esclusivamente come sfondo cartografico.",
        "Non devono essere rimosse le attribuzioni mostrate sulla mappa.",
        "Lo sfondo Esri non modifica la licenza dei dati del censimento visualizzati sopra la mappa."
      ]
    }
  ]);

  const MODULES = Object.freeze({
    trees: {
      pageId: "tree-search-page",
      headerSelector: ".tree-search-header",
      icon: "🌳",
      title: "Catasto alberi",
      subtitle: "Comune di Bologna · Alberi in manutenzione",
      summary: "La ricerca e i punti sulla mappa provengono dal dataset ufficiale del Comune di Bologna. Varga Cantieri filtra, ordina e rappresenta i record senza dichiararsi servizio ufficiale del Comune.",
      sources: [
        {
          name: "Comune di Bologna – Alberi in manutenzione",
          role: "Numero punto, codice albero, specie, caratteristiche e coordinate",
          license: "Creative Commons Attribuzione 4.0 Internazionale (CC BY 4.0)",
          sourceUrl: "https://opendata.comune.bologna.it/explore/dataset/alberi-manutenzioni/",
          apiUrl: "https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets/alberi-manutenzioni/records",
          licenseUrl: "https://creativecommons.org/licenses/by/4.0/deed.it",
          termsUrl: "https://opendata.comune.bologna.it/terms/terms-and-conditions/",
          attribution: "Fonte: Comune di Bologna – Portale Open Data, dataset “Alberi in manutenzione”, CC BY 4.0. Dati filtrati e rielaborati da Varga Cantieri.",
          conditions: [
            "È consentito consultare, copiare, condividere e adattare i dati, anche per finalità commerciali.",
            "Occorre indicare il Comune di Bologna come fonte, collegare la licenza e segnalare le rielaborazioni effettuate.",
            "L’uso dei dati non implica approvazione, patrocinio o responsabilità del Comune nei confronti di Varga Cantieri."
          ]
        },
        ...SHARED_MAP_SOURCES
      ]
    },
    greenAreas: {
      pageId: "green-areas-page",
      headerSelector: ".green-areas-header",
      icon: "🌿",
      title: "Aree verdi",
      subtitle: "DBTR Regione Emilia-Romagna + OpenStreetMap",
      summary: "Il livello ufficiale delle aree verdi è fornito dal Database Topografico Regionale. Nomi, categorie e risultati ricercabili possono essere integrati con dati OpenStreetMap interrogati tramite Overpass.",
      sources: [
        {
          name: "Regione Emilia-Romagna – Database Topografico Regionale (DBTR)",
          role: "Livello cartografico ufficiale PSR_Area_verde tramite servizio WMS",
          license: "Creative Commons Attribuzione 4.0 Internazionale (CC BY 4.0)",
          sourceUrl: "https://geoportale.regione.emilia-romagna.it/approfondimenti/database-topografico-regionale",
          apiUrl: "https://servizigis.regione.emilia-romagna.it/wms/dbtr",
          licenseUrl: "https://creativecommons.org/licenses/by/4.0/deed.it",
          attribution: "Fonte: Regione Emilia-Romagna – Database Topografico Regionale, CC BY 4.0. Visualizzazione e filtri: Varga Cantieri.",
          conditions: [
            "È consentito riutilizzare e adattare i dati, anche per finalità commerciali, mantenendo l’attribuzione alla Regione Emilia-Romagna.",
            "Occorre indicare la licenza CC BY 4.0 e dichiarare eventuali modifiche o rielaborazioni.",
            "Il WMS è un servizio esterno: disponibilità, copertura e aggiornamento dipendono dall’ente titolare."
          ]
        },
        SHARED_MAP_SOURCES[0],
        SHARED_MAP_SOURCES[1]
      ]
    },
    urbanFurniture: {
      pageId: "urban-furniture-page",
      headerSelector: ".urban-furniture-header",
      icon: "🪑",
      title: "Arredo urbano",
      subtitle: "OpenStreetMap · interrogazione Overpass",
      summary: "Panchine, cestini, fontane, aree giochi, idranti e le altre categorie sono ricavate dai tag presenti in OpenStreetMap. Il servizio interno dell’app è soltanto un tramite tecnico verso i dati OSM e non ne cambia la licenza.",
      sources: [
        {
          ...SHARED_MAP_SOURCES[0],
          role: "Posizione, categoria e attributi degli elementi di arredo urbano",
          attribution: "© contributori OpenStreetMap · dati disponibili con licenza ODbL 1.0",
          conditions: [
            "È consentito usare, copiare e adattare i dati, anche commercialmente, con attribuzione a OpenStreetMap e ai contributori.",
            "La licenza ODbL e l’attribuzione devono restare chiaramente accessibili.",
            "I dati sono inseriti dalla comunità e possono essere incompleti, duplicati, non aggiornati o privi di alcuni dettagli.",
            "Una banca dati derivata distribuita pubblicamente può richiedere la condivisione alle stesse condizioni ODbL."
          ]
        },
        SHARED_MAP_SOURCES[1]
      ]
    },
    wastewater: {
      pageId: "wastewater-plants-page",
      headerSelector: ".wastewater-plants-header",
      icon: "🏭",
      title: "Censimento depuratori",
      subtitle: "ARPAE Emilia-Romagna + OpenStreetMap",
      summary: "I depuratori provengono dal servizio geografico ufficiale ARPAE. I sollevamenti e le stazioni di pompaggio integrative provengono da OpenStreetMap e sono sempre identificati separatamente nell’app.",
      sources: [
        {
          name: "ARPAE Emilia-Romagna – Depuratori della Regione Emilia-Romagna, edizione 2023",
          role: "Impianti di trattamento delle acque reflue urbane, dati tecnici e coordinate",
          license: "Creative Commons Attribuzione 4.0 Internazionale (CC BY 4.0)",
          sourceUrl: "https://dati.arpae.it/dataset/arpa_acq_reflue_urbane_depurate_depurat_tutti_22_e23",
          apiUrl: "https://servizi-gis.arpae.it/server/rest/services/Geoportal/ACQUEPressioni/MapServer/1/query",
          licenseUrl: "https://creativecommons.org/licenses/by/4.0/deed.it",
          attribution: "Fonte: ARPAE Emilia-Romagna – “Depuratori della Regione Emilia-Romagna, edizione 2023”, CC BY 4.0. Dati filtrati e rappresentati da Varga Cantieri.",
          conditions: [
            "È consentito riutilizzare, condividere e adattare i dati, anche per finalità commerciali.",
            "Occorre attribuire ARPAE Emilia-Romagna, collegare la licenza CC BY 4.0 e dichiarare le rielaborazioni.",
            "La scheda dell’app non sostituisce atti autorizzativi, dati gestionali del titolare dell’impianto o verifiche tecniche sul posto."
          ]
        },
        {
          ...SHARED_MAP_SOURCES[0],
          role: "Sollevamenti fognari e stazioni di pompaggio integrative",
          attribution: "© contributori OpenStreetMap · sollevamenti e stazioni di pompaggio sotto licenza ODbL 1.0",
          conditions: [
            "I punti OSM sono distinti dai depuratori ARPAE e non devono essere considerati un censimento istituzionale completo.",
            "Occorre mantenere attribuzione e riferimento alla licenza ODbL.",
            "La classificazione dipende dai tag inseriti dai contributori e deve essere verificata prima di impieghi tecnici o operativi."
          ]
        },
        SHARED_MAP_SOURCES[1]
      ]
    }
  });

  let activeModuleKey = "";
  let sourcePage = null;
  let sourceScrollY = 0;

  const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  }[char]));

  function externalLink(url, label, className = "data-license-link") {
    if (!url) return "";
    return `<a class="${className}" href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)} ↗</a>`;
  }

  function injectStyle() {
    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
      .data-license-trigger {
        margin-left: auto;
        min-height: 42px;
        padding: 10px 14px;
        border: 1px solid rgba(255,255,255,.38);
        border-radius: 14px;
        background: rgba(8,38,68,.58);
        color: #fff;
        font-weight: 800;
        letter-spacing: .02em;
        white-space: nowrap;
        box-shadow: 0 8px 24px rgba(0,0,0,.16);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
      }
      .data-license-trigger:hover,
      .data-license-trigger:focus-visible { background: rgba(15,93,159,.9); transform: translateY(-1px); }
      .data-license-page {
        position: relative;
        z-index: 1100;
        min-height: 100vh;
        padding: max(14px, env(safe-area-inset-top)) max(14px, env(safe-area-inset-right)) max(24px, env(safe-area-inset-bottom)) max(14px, env(safe-area-inset-left));
        color: #eff8ff;
        background:
          radial-gradient(circle at top right, rgba(22,135,205,.28), transparent 34rem),
          linear-gradient(155deg, #071522 0%, #0a2033 48%, #07131f 100%);
        overflow-x: hidden;
      }
      .data-license-page.hidden { display: none !important; }
      .data-license-shell { width: min(1040px, 100%); margin: 0 auto; }
      .data-license-header {
        display: grid;
        grid-template-columns: auto minmax(0, 1fr);
        gap: 14px;
        align-items: center;
        margin-bottom: 18px;
        padding: 15px;
        border: 1px solid rgba(151,209,245,.25);
        border-radius: 22px;
        background: rgba(9,33,53,.78);
        box-shadow: 0 18px 50px rgba(0,0,0,.24);
        backdrop-filter: blur(18px);
        -webkit-backdrop-filter: blur(18px);
      }
      .data-license-back { min-height: 44px; border-radius: 14px; }
      .data-license-kicker { display: block; margin-bottom: 3px; color: #8ed4ff; font-size: .78rem; font-weight: 900; letter-spacing: .1em; text-transform: uppercase; }
      .data-license-header h1 { margin: 0; color: #fff; font-size: clamp(1.35rem, 4vw, 2rem); line-height: 1.1; }
      .data-license-header p { margin: 6px 0 0; color: #b9d8eb; }
      .data-license-intro,
      .data-license-source-card,
      .data-license-general {
        border: 1px solid rgba(151,209,245,.2);
        border-radius: 20px;
        background: rgba(11,39,61,.72);
        box-shadow: 0 12px 36px rgba(0,0,0,.17);
      }
      .data-license-intro { display: grid; grid-template-columns: auto 1fr; gap: 14px; padding: 18px; margin-bottom: 16px; }
      .data-license-intro-icon { display: grid; place-items: center; width: 54px; height: 54px; border-radius: 17px; background: rgba(44,157,225,.2); font-size: 1.75rem; }
      .data-license-intro h2 { margin: 0 0 5px; color: #fff; }
      .data-license-intro p { margin: 0; color: #c6dfed; line-height: 1.55; }
      .data-license-source-list { display: grid; gap: 14px; }
      .data-license-source-card { padding: 18px; }
      .data-license-source-head { display: flex; gap: 12px; justify-content: space-between; align-items: flex-start; }
      .data-license-source-card h2 { margin: 0; color: #fff; font-size: 1.08rem; line-height: 1.3; }
      .data-license-role { margin: 6px 0 0; color: #b6d7e9; line-height: 1.45; }
      .data-license-badge { flex: 0 0 auto; padding: 6px 9px; border: 1px solid rgba(101,213,156,.42); border-radius: 999px; color: #a9f0ca; background: rgba(19,115,71,.22); font-size: .72rem; font-weight: 900; }
      .data-license-grid { display: grid; grid-template-columns: minmax(0, 1fr) minmax(0, 1fr); gap: 12px; margin-top: 15px; }
      .data-license-detail { padding: 13px; border-radius: 15px; background: rgba(2,18,30,.42); }
      .data-license-detail span { display: block; color: #83c9f4; font-size: .75rem; font-weight: 900; text-transform: uppercase; letter-spacing: .06em; }
      .data-license-detail strong,
      .data-license-detail code { display: block; margin-top: 5px; color: #f4fbff; overflow-wrap: anywhere; line-height: 1.45; }
      .data-license-detail code { font-size: .78rem; font-weight: 600; }
      .data-license-conditions { margin: 15px 0 0; padding-left: 1.25rem; color: #d6e9f3; line-height: 1.55; }
      .data-license-conditions li + li { margin-top: 7px; }
      .data-license-links { display: flex; flex-wrap: wrap; gap: 9px; margin-top: 15px; }
      .data-license-link { display: inline-flex; align-items: center; min-height: 38px; padding: 8px 11px; border: 1px solid rgba(116,195,242,.36); border-radius: 12px; color: #bfe7ff; background: rgba(13,83,126,.28); font-weight: 800; text-decoration: none; }
      .data-license-link:hover,
      .data-license-link:focus-visible { color: #fff; background: rgba(20,113,169,.58); }
      .data-license-general { margin-top: 16px; padding: 18px; }
      .data-license-general h2 { margin: 0 0 10px; color: #fff; font-size: 1.08rem; }
      .data-license-general p { margin: 0; color: #c9dfeb; line-height: 1.55; }
      .data-license-general p + p { margin-top: 10px; }
      .data-license-verified { margin: 16px 2px 0; color: #8db8cf; font-size: .82rem; text-align: center; }
      @media (max-width: 760px) {
        .wastewater-plants-header,
        .tree-search-header,
        .green-areas-header,
        .urban-furniture-header { flex-wrap: wrap; }
        .data-license-trigger { width: 100%; margin-left: 0; }
        .data-license-grid { grid-template-columns: 1fr; }
        .data-license-source-head { display: grid; }
        .data-license-badge { justify-self: start; }
      }
      @media (max-width: 520px) {
        .data-license-header { grid-template-columns: 1fr; }
        .data-license-back { width: 100%; }
        .data-license-intro { grid-template-columns: 1fr; }
      }
    `;
    document.head.appendChild(style);
  }

  function renderSource(source) {
    const links = [
      externalLink(source.sourceUrl, "Apri la fonte ufficiale"),
      externalLink(source.apiUrl, "Apri API / servizio"),
      externalLink(source.licenseUrl, "Leggi la licenza"),
      externalLink(source.termsUrl, "Condizioni del portale")
    ].filter(Boolean).join("");
    return `
      <article class="data-license-source-card">
        <div class="data-license-source-head">
          <div>
            <h2>${escapeHtml(source.name)}</h2>
            <p class="data-license-role">${escapeHtml(source.role)}</p>
          </div>
          <span class="data-license-badge">FONTE ESTERNA</span>
        </div>
        <div class="data-license-grid">
          <div class="data-license-detail"><span>Licenza / termini</span><strong>${escapeHtml(source.license)}</strong></div>
          <div class="data-license-detail"><span>Attribuzione da mantenere</span><strong>${escapeHtml(source.attribution)}</strong></div>
          ${source.apiUrl ? `<div class="data-license-detail"><span>Endpoint utilizzato</span><code>${escapeHtml(source.apiUrl)}</code></div>` : ""}
          <div class="data-license-detail"><span>Rielaborazione nell’app</span><strong>Ricerca, filtri, normalizzazione, cache temporanea e rappresentazione cartografica.</strong></div>
        </div>
        <ul class="data-license-conditions">${source.conditions.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
        <div class="data-license-links">${links}</div>
      </article>
    `;
  }

  function buildView() {
    if (document.getElementById(VIEW_ID)) return document.getElementById(VIEW_ID);
    const view = document.createElement("section");
    view.id = VIEW_ID;
    view.className = "data-license-page hidden";
    view.setAttribute("aria-hidden", "true");
    view.setAttribute("aria-labelledby", "data-license-title");
    view.innerHTML = `
      <div class="data-license-shell">
        <header class="data-license-header">
          <button id="data-license-back-btn" class="btn data-license-back" type="button">← TORNA AL MODULO</button>
          <div>
            <span class="data-license-kicker">Trasparenza dei dati</span>
            <h1 id="data-license-title">Licenze e fonti dati</h1>
            <p id="data-license-subtitle"></p>
          </div>
        </header>
        <main id="data-license-content"></main>
      </div>
    `;
    document.body.appendChild(view);
    view.querySelector("#data-license-back-btn")?.addEventListener("click", closeView);
    return view;
  }

  function renderModule(moduleKey) {
    const module = MODULES[moduleKey];
    const view = buildView();
    const subtitle = view.querySelector("#data-license-subtitle");
    const content = view.querySelector("#data-license-content");
    if (!module || !subtitle || !content) return;
    subtitle.textContent = `${module.icon} ${module.title} · ${module.subtitle}`;
    content.innerHTML = `
      <section class="data-license-intro">
        <div class="data-license-intro-icon" aria-hidden="true">${module.icon}</div>
        <div><h2>${escapeHtml(module.title)}</h2><p>${escapeHtml(module.summary)}</p></div>
      </section>
      <section class="data-license-source-list" aria-label="Fonti e licenze">${module.sources.map(renderSource).join("")}</section>
      <section class="data-license-general">
        <h2>Condizioni generali e responsabilità</h2>
        <p>Queste informazioni spiegano la provenienza dei dati mostrati da Varga Cantieri; non trasferiscono a Varga Cantieri la titolarità dei dataset e non sostituiscono i testi legali collegati.</p>
        <p>I dati possono essere incompleti, non aggiornati o temporaneamente indisponibili. Prima di lavori, accessi, scavi, interventi di sicurezza o decisioni tecniche è necessario verificare le informazioni presso l’ente, il gestore o il proprietario competente.</p>
        <p>Varga Cantieri non è un’app ufficiale del Comune di Bologna, della Regione Emilia-Romagna, di ARPAE, di OpenStreetMap Foundation o di Esri. I nomi degli enti sono indicati esclusivamente per attribuire correttamente le fonti.</p>
      </section>
      <p class="data-license-verified">Informazioni sulle licenze verificate il ${VERIFIED_ON}. In caso di contrasto prevalgono sempre i termini pubblicati dalla fonte ufficiale.</p>
    `;
  }

  function openView(moduleKey) {
    const module = MODULES[moduleKey];
    if (!module) return;
    const page = document.getElementById(module.pageId);
    const view = buildView();
    activeModuleKey = moduleKey;
    sourcePage = page || null;
    sourceScrollY = window.scrollY || 0;
    renderModule(moduleKey);
    sourcePage?.classList.add("hidden");
    sourcePage?.setAttribute("aria-hidden", "true");
    view.classList.remove("hidden");
    view.setAttribute("aria-hidden", "false");
    document.body.classList.add("data-license-view-open");
    window.scrollTo({ top: 0, behavior: "auto" });
    window.setTimeout(() => view.querySelector("#data-license-back-btn")?.focus(), 0);
  }

  function closeView() {
    const view = document.getElementById(VIEW_ID);
    if (!view || view.classList.contains("hidden")) return;
    view.classList.add("hidden");
    view.setAttribute("aria-hidden", "true");
    document.body.classList.remove("data-license-view-open");
    sourcePage?.classList.remove("hidden");
    sourcePage?.setAttribute("aria-hidden", "false");
    const returnButton = sourcePage?.querySelector?.(`[data-license-module="${activeModuleKey}"]`);
    window.scrollTo({ top: sourceScrollY, behavior: "auto" });
    window.setTimeout(() => returnButton?.focus(), 0);
    activeModuleKey = "";
    sourcePage = null;
  }

  function addModuleButton(moduleKey, module) {
    const page = document.getElementById(module.pageId);
    if (!page || page.querySelector(`[data-license-module="${moduleKey}"]`)) return;
    const header = page.querySelector(module.headerSelector) || page.querySelector("header") || page.firstElementChild;
    if (!header) return;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "btn data-license-trigger";
    button.dataset.licenseModule = moduleKey;
    button.setAttribute("aria-label", `Apri licenze e fonti dati di ${module.title}`);
    button.innerHTML = "⚖️ LICENZA";
    button.addEventListener("click", () => openView(moduleKey));
    header.appendChild(button);
  }

  function install() {
    injectStyle();
    buildView();
    Object.entries(MODULES).forEach(([moduleKey, module]) => addModuleButton(moduleKey, module));

    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape" && !document.getElementById(VIEW_ID)?.classList.contains("hidden")) closeView();
    });

    window.HeraDataLicenses = Object.freeze({
      open: openView,
      close: closeView,
      modules: Object.keys(MODULES)
    });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();

(() => {
  "use strict";

  const PAGE_ID = "verde-bologna-page";
  const STYLE_ID = "verde-bologna-operativo-style";
  const SELECT_ID = "verde-bologna-operativo-category";
  const MOBILE_QUERY = "(max-width: 760px)";
  const DATASETS = [
    ["un_gest", "🌳 Aree verdi in manutenzione"],
    ["alberi-manutenzioni", "🌲 Alberi singoli"],
    ["popolazione-arborea", "🌴 Popolazioni arboree"],
    ["siepi", "🌿 Siepi in manutenzione"],
    ["attrezzature_ludiche_ginniche_sportive", "🛝 Giochi e attrezzature"],
    ["arredo", "🪑 Arredo urbano comunale"],
    ["sgambatura_cani", "🐕 Aree cani"],
    ["carta-tecnica-comunale-toponimi-parchi-e-giardini", "🏞️ Parchi e giardini"],
    ["aree-verdi_entrate_centroidi", "🚪 Ingressi aree verdi"],
    ["aree-ortive", "🥕 Aree ortive"],
    ["verde_privato_urbanizzato", "🏡 Verde privato"]
  ];

  const $ = (id) => document.getElementById(id);

  function injectStyle() {
    if ($(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
      .verde-bologna-operativo-card{display:none}
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
    });
  }

  function syncFromCards(page) {
    const select = $(SELECT_ID);
    if (!select || !page) return;
    page.addEventListener("click", (event) => {
      const button = event.target?.closest?.("[data-vb-open]");
      const id = button?.getAttribute("data-vb-open");
      if (id && [...select.options].some((option) => option.value === id)) select.value = id;
    }, true);
  }

  function primeDefaultDataset(page) {
    if (!page || !window.matchMedia(MOBILE_QUERY).matches) return;
    const browser = $("verde-bologna-browser");
    if (browser && !browser.classList.contains("hidden")) return;
    page.querySelector('[data-vb-open="un_gest"]')?.click();
  }

  function observePage(page) {
    if (!page || page.dataset.operativoObserved === "1") return;
    page.dataset.operativoObserved = "1";
    const observer = new MutationObserver(() => {
      if (!page.classList.contains("hidden")) {
        ensureOperationalCard(page);
        window.setTimeout(() => primeDefaultDataset(page), 40);
      }
    });
    observer.observe(page, { attributes: true, attributeFilter: ["class", "aria-hidden"] });
  }

  function install() {
    injectStyle();
    const attach = () => {
      const page = $(PAGE_ID);
      if (!page) return false;
      ensureOperationalCard(page);
      syncFromCards(page);
      observePage(page);
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

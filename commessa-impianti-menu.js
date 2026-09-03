/* Accesso rapido alla gestione commessa dalla testata, ottimizzato per telefono. */
(() => {
  "use strict";

  function ensureAccountingRoundsModule() {
    if (window.VargaAccountingRounds || document.querySelector('script[data-varga-accounting-rounds]')) return;
    const script = document.createElement("script");
    script.src = `accounting-rounds.js?v=20260903-2`;
    script.async = true;
    script.dataset.vargaAccountingRounds = "1";
    document.head.appendChild(script);
  }

  ensureAccountingRoundsModule();

  let initialized = false;
  let availabilityTimer = null;

  function ensureMobileView() {
    if (document.getElementById("commessa-mobile-management")) return;
    const panel = document.getElementById("panel-commesse");
    if (!panel) return;
    panel.insertAdjacentHTML("beforeend", `
      <section id="commessa-mobile-management" class="commessa-mobile-management hidden" aria-hidden="true">
        <header class="commessa-mobile-management-head">
          <button id="commessa-mobile-management-back" class="btn" type="button">← Commessa</button>
          <div><p class="management-eyebrow">GESTIONE COMMESSA</p><h2 id="commessa-mobile-management-title">Gestione commessa</h2><p id="commessa-mobile-management-meta" class="muted"></p></div>
        </header>
        <div id="commessa-mobile-management-stats" class="commessa-mobile-management-stats"></div>
        <section id="commessa-mobile-management-home">
          <div class="commessa-mobile-action-grid">
            <button type="button" data-commessa-mobile-action="add"><span aria-hidden="true">＋</span><strong>Aggiungi impianto</strong><small>Compila tutti i dati con un modulo guidato</small></button>
            <button type="button" data-commessa-mobile-action="edit"><span aria-hidden="true">✏️</span><strong>Modifica impianti</strong><small>Cerca un impianto e modifica ogni campo</small></button>
            <button type="button" data-commessa-mobile-action="import"><span aria-hidden="true">📥</span><strong>Importa impianti</strong><small>Carica Excel, CSV o Google Sheet</small></button>
            <button type="button" data-commessa-mobile-action="export"><span aria-hidden="true">📤</span><strong>Esporta impianti</strong><small>Scarica matrice e contabilità</small></button>
            <button type="button" data-commessa-mobile-action="prices"><span aria-hidden="true">€</span><strong>Prezziario</strong><small>Gestisci prezzi, ribassi e voci</small></button>
            <button type="button" data-commessa-mobile-action="advanced"><span aria-hidden="true">▦</span><strong>Vista completa</strong><small>Apri la tabella amministrativa</small></button>
          </div>
        </section>
        <section id="commessa-mobile-plant-list" class="hidden" aria-hidden="true">
          <div class="commessa-mobile-section-head"><button type="button" class="btn" data-commessa-mobile-action="home">← Gestione</button><div><h3>Modifica impianti</h3><p class="muted">Cerca e seleziona l’impianto da modificare.</p></div></div>
          <label class="commessa-mobile-search"><span aria-hidden="true">⌕</span><input id="commessa-mobile-plant-search" type="search" placeholder="Nome, comune, ID SAP o codice prezzo" autocomplete="off"></label>
          <div id="commessa-mobile-plant-results" class="commessa-mobile-plant-results"></div>
        </section>
        <section id="commessa-mobile-plant-editor" class="hidden" aria-hidden="true">
          <div class="commessa-mobile-section-head"><button id="commessa-mobile-editor-back" type="button" class="btn">← Indietro</button><div><p class="management-eyebrow" id="commessa-mobile-editor-eyebrow">NUOVO IMPIANTO</p><h3 id="commessa-mobile-editor-title">Aggiungi impianto</h3><p class="muted">I campi economici vengono calcolati dal prezziario.</p></div></div>
          <form id="commessa-mobile-plant-form" class="commessa-mobile-plant-form"><div id="commessa-mobile-plant-fields"></div><button id="commessa-mobile-add-work" class="btn commessa-mobile-add-work hidden" type="button">＋ Nuova lavorazione sullo stesso impianto</button><p id="commessa-mobile-plant-feedback" class="row-feedback" role="alert"></p><div class="commessa-mobile-form-actions"><button id="commessa-mobile-plant-cancel" class="btn" type="button">Annulla</button><button id="commessa-mobile-plant-save" class="btn btn-primary" type="submit">Salva impianto</button></div></form>
        </section>
      </section>`);
  }

  function currentElements() {
    return {
      wrap: document.getElementById("commessa-plants-menu-wrap"),
      toggle: document.getElementById("commessa-plants-menu-btn"),
      title: document.getElementById("commessa-focus-label")
    };
  }

  function updateAvailability() {
    const { wrap } = currentElements();
    if (!wrap) return;
    const hasCommessa = typeof selectedCommessaId !== "undefined" && Boolean(String(selectedCommessaId || "").trim());
    const allowed = typeof canManageData === "function" && canManageData();
    wrap.classList.toggle("hidden", !(hasCommessa && allowed));
  }

  async function openCommessaManagement(toggle) {
    if (typeof canManageData !== "function" || !canManageData()) {
      window.alert("La gestione della commessa è riservata agli amministratori.");
      return;
    }
    const commessaId = typeof selectedCommessaId === "undefined" ? "" : String(selectedCommessaId || "").trim();
    const commessa = commessaId && typeof commesseById !== "undefined" ? commesseById.get(commessaId) : null;
    if (!commessa) {
      window.alert("Seleziona prima una commessa.");
      return;
    }
    if (!window.AccountingV2?.openMobileHub) {
      window.alert("La gestione commessa non è ancora disponibile. Ricarica l’app e riprova.");
      return;
    }
    toggle.disabled = true;
    try {
      ensureMobileView();
      openManagementPanel("commesse");
      await window.AccountingV2.openMobileHub(commessa);
    } finally {
      toggle.disabled = false;
    }
  }

  function initialize() {
    if (initialized) {
      updateAvailability();
      return true;
    }
    const { wrap, toggle, title } = currentElements();
    if (!wrap || !toggle || !title) return false;

    initialized = true;
    ensureMobileView();
    document.getElementById("commessa-plants-menu")?.remove();
    toggle.removeAttribute("aria-haspopup");
    toggle.removeAttribute("aria-expanded");
    toggle.removeAttribute("aria-controls");
    toggle.setAttribute("aria-label", "Apri gestione commessa");
    toggle.setAttribute("title", "Gestione commessa");
    toggle.addEventListener("click", event => {
      event.preventDefault();
      void openCommessaManagement(toggle);
    });

    new MutationObserver(updateAvailability).observe(title, { childList: true, characterData: true, subtree: true });
    updateAvailability();
    return true;
  }

  function bootUntilReady() {
    if (initialize()) return;
    let attempts = 0;
    const timer = setInterval(() => {
      attempts += 1;
      if (initialize() || attempts >= 80) clearInterval(timer);
    }, 250);
  }

  window.addEventListener("pageshow", () => { initialize(); updateAvailability(); });
  window.addEventListener("focus", updateAvailability);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) updateAvailability(); });
  document.addEventListener("click", event => {
    if (event.target.closest?.('[data-commessa-id], .commessa-card, .commessa-item, [data-action="open-commessa"]')) {
      setTimeout(updateAvailability, 0);
      setTimeout(updateAvailability, 150);
    }
  }, true);

  availabilityTimer = setInterval(updateAvailability, 1000);
  window.addEventListener("beforeunload", () => clearInterval(availabilityTimer), { once: true });

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", bootUntilReady, { once: true });
  else bootUntilReady();
})();

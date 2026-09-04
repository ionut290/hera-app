/* Menu tre puntini commessa: ricostruito da zero, indipendente e mobile-first. */
(() => {
  "use strict";

  const IDS = {
    wrap: "commessa-plants-menu-wrap",
    toggle: "commessa-plants-menu-btn",
    menu: "commessa-plants-menu",
    panel: "panel-commesse",
    mobile: "commessa-mobile-management",
    mobileHome: "commessa-mobile-management-home",
    mobileList: "commessa-mobile-plant-list",
    mobileEditor: "commessa-mobile-plant-editor",
    mobileForm: "commessa-mobile-plant-form"
  };

  let busy = false;
  let editorLockUntil = 0;

  function el(id) {
    return document.getElementById(id);
  }

  function canOpen() {
    try {
      return typeof canManageData === "function"
        && canManageData()
        && typeof selectedCommessaId !== "undefined"
        && Boolean(String(selectedCommessaId || "").trim());
    } catch (_) {
      return false;
    }
  }

  function selectedCommessa() {
    try {
      const id = String(selectedCommessaId || "").trim();
      return id && typeof commesseById !== "undefined" ? commesseById.get(id) : null;
    } catch (_) {
      return null;
    }
  }

  function ensureMobileView() {
    if (el(IDS.mobile)) return true;
    const panel = el(IDS.panel);
    if (!panel) return false;

    panel.insertAdjacentHTML("beforeend", `
      <section id="commessa-mobile-management" class="commessa-mobile-management hidden" aria-hidden="true">
        <header class="commessa-mobile-management-head">
          <button id="commessa-mobile-management-back" class="btn" type="button">← Commessa</button>
          <div>
            <p class="management-eyebrow">GESTIONE COMMESSA</p>
            <h2 id="commessa-mobile-management-title">Gestione commessa</h2>
            <p id="commessa-mobile-management-meta" class="muted"></p>
          </div>
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
          <div class="commessa-mobile-section-head">
            <button type="button" class="btn" data-commessa-mobile-action="home">← Gestione</button>
            <div><h3>Modifica impianti</h3><p class="muted">Cerca e seleziona l’impianto da modificare.</p></div>
          </div>
          <label class="commessa-mobile-search"><span aria-hidden="true">⌕</span><input id="commessa-mobile-plant-search" type="search" placeholder="Nome, comune, ID SAP o codice prezzo" autocomplete="off"></label>
          <div id="commessa-mobile-plant-results" class="commessa-mobile-plant-results"></div>
        </section>
        <section id="commessa-mobile-plant-editor" class="hidden" aria-hidden="true">
          <div class="commessa-mobile-section-head">
            <button id="commessa-mobile-editor-back" type="button" class="btn">← Indietro</button>
            <div><p class="management-eyebrow" id="commessa-mobile-editor-eyebrow">NUOVO IMPIANTO</p><h3 id="commessa-mobile-editor-title">Aggiungi impianto</h3><p class="muted">I campi economici vengono calcolati dal prezziario.</p></div>
          </div>
          <form id="commessa-mobile-plant-form" class="commessa-mobile-plant-form">
            <div id="commessa-mobile-plant-fields"></div>
            <button id="commessa-mobile-add-work" class="btn commessa-mobile-add-work hidden" type="button">＋ Nuova lavorazione sullo stesso impianto</button>
            <p id="commessa-mobile-plant-feedback" class="row-feedback" role="alert"></p>
            <div class="commessa-mobile-form-actions">
              <button id="commessa-mobile-plant-cancel" class="btn" type="button">Annulla</button>
              <button id="commessa-mobile-plant-save" class="btn btn-primary" type="submit">Salva impianto</button>
            </div>
          </form>
        </section>
      </section>`);
    return true;
  }

  function setMenuOpen(open) {
    const toggle = el(IDS.toggle);
    const menu = el(IDS.menu);
    if (!toggle || !menu) return;
    menu.classList.toggle("hidden", !open);
    menu.setAttribute("aria-hidden", open ? "false" : "true");
    toggle.setAttribute("aria-expanded", open ? "true" : "false");
  }

  function closeMenu() {
    setMenuOpen(false);
  }

  function updateAvailability() {
    const wrap = el(IDS.wrap);
    if (!wrap) return;
    const available = canOpen();
    wrap.classList.toggle("hidden", !available);
    if (!available) closeMenu();
  }

  function restoreEditorIfLocked() {
    if (Date.now() > editorLockUntil) return;
    const mobile = el(IDS.mobile);
    const home = el(IDS.mobileHome);
    const list = el(IDS.mobileList);
    const editor = el(IDS.mobileEditor);
    const form = el(IDS.mobileForm);
    if (!mobile || !editor || !form) return;
    mobile.classList.remove("hidden");
    mobile.setAttribute("aria-hidden", "false");
    home?.classList.add("hidden");
    home?.setAttribute("aria-hidden", "true");
    list?.classList.add("hidden");
    list?.setAttribute("aria-hidden", "true");
    editor.classList.remove("hidden");
    editor.setAttribute("aria-hidden", "false");
  }

  async function openAction(action) {
    if (busy) return;
    if (!canOpen()) {
      window.alert("Seleziona prima una commessa.");
      return;
    }
    const commessa = selectedCommessa();
    if (!commessa) {
      window.alert("Commessa non disponibile. Riaprila e riprova.");
      return;
    }
    if (!ensureMobileView()) {
      window.alert("Gestione commessa non disponibile. Ricarica l’app e riprova.");
      return;
    }
    if (!window.AccountingV2?.openMobileHub) {
      window.alert("Modulo gestione impianti non ancora caricato. Ricarica l’app e riprova.");
      return;
    }

    busy = true;
    const toggle = el(IDS.toggle);
    if (toggle) toggle.disabled = true;
    closeMenu();

    try {
      if (typeof openManagementPanel === "function") openManagementPanel("commesse");
      const opened = await window.AccountingV2.openMobileHub(commessa);
      if (!opened) throw new Error("Apertura interrotta");

      if (!action || action === "home") return;
      const target = document.querySelector(`[data-commessa-mobile-action="${CSS.escape(action)}"]`);
      if (!target) throw new Error(`Azione ${action} non disponibile`);
      target.click();
    } catch (error) {
      console.error("Menu tre puntini: apertura azione non riuscita", error);
      window.alert("Non è stato possibile aprire questa funzione. Riprova.");
    } finally {
      busy = false;
      if (toggle) toggle.disabled = false;
    }
  }

  function handleToggle(event) {
    const toggle = event.target.closest?.(`#${IDS.toggle}`);
    if (!toggle) return false;
    event.preventDefault();
    event.stopPropagation();
    if (!canOpen()) return true;
    const menu = el(IDS.menu);
    setMenuOpen(Boolean(menu?.classList.contains("hidden")));
    return true;
  }

  function handleMenuAction(event) {
    const button = event.target.closest?.("[data-commessa-plants-action]");
    if (!button) return false;
    event.preventDefault();
    event.stopPropagation();
    void openAction(button.dataset.commessaPlantsAction || "home");
    return true;
  }

  function handleDocumentClick(event) {
    if (handleToggle(event)) return;
    if (handleMenuAction(event)) return;
    if (!event.target.closest?.(`#${IDS.wrap}`)) closeMenu();

    if (event.target.closest?.("#commessa-mobile-current-location")) {
      editorLockUntil = Date.now() + 20000;
      setTimeout(restoreEditorIfLocked, 0);
    }
  }

  function initialize() {
    const wrap = el(IDS.wrap);
    const toggle = el(IDS.toggle);
    const menu = el(IDS.menu);
    if (!wrap || !toggle || !menu) return false;

    ensureMobileView();
    toggle.setAttribute("aria-label", "Apri menu gestione commessa");
    toggle.setAttribute("title", "Gestione commessa");
    toggle.setAttribute("aria-haspopup", "menu");
    toggle.setAttribute("aria-controls", IDS.menu);
    setMenuOpen(false);
    updateAvailability();
    return true;
  }

  document.addEventListener("click", handleDocumentClick, false);
  document.addEventListener("keydown", event => {
    if (event.key === "Escape") closeMenu();
  });
  window.addEventListener("pageshow", () => {
    initialize();
    restoreEditorIfLocked();
  });
  window.addEventListener("focus", () => {
    updateAvailability();
    restoreEditorIfLocked();
  });
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      updateAvailability();
      restoreEditorIfLocked();
    }
  });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialize, { once: true });
  } else {
    initialize();
  }

  const observer = new MutationObserver(() => updateAvailability());
  const title = el("commessa-focus-label");
  if (title) observer.observe(title, { childList: true, characterData: true, subtree: true });

  window.CommessaThreeDotsMenu = {
    version: "2026.09.04-rebuild1",
    open: openAction,
    refresh: updateAvailability
  };
})();

/* Accesso rapido alla gestione commessa dalla testata, ottimizzato per telefono. */
(() => {
  "use strict";

  const wrap = document.getElementById("commessa-plants-menu-wrap");
  const toggle = document.getElementById("commessa-plants-menu-btn");
  const title = document.getElementById("commessa-focus-label");
  if (!wrap || !toggle || !title) return;

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

  function restoreMobilePlantEditor(form) {
    if (!form || !document.contains(form)) return false;
    const mobile = document.getElementById("commessa-mobile-management");
    const home = document.getElementById("commessa-mobile-management-home");
    const list = document.getElementById("commessa-mobile-plant-list");
    const editor = document.getElementById("commessa-mobile-plant-editor");
    if (!mobile || !editor || !editor.contains(form)) return false;

    mobile.classList.remove("hidden");
    mobile.setAttribute("aria-hidden", "false");
    home?.classList.add("hidden");
    home?.setAttribute("aria-hidden", "true");
    list?.classList.add("hidden");
    list?.setAttribute("aria-hidden", "true");
    editor.classList.remove("hidden");
    editor.setAttribute("aria-hidden", "false");
    return true;
  }

  function handleMobileCurrentLocation(event) {
    const button = event.target.closest?.("#commessa-mobile-current-location");
    if (!button) return;

    event.preventDefault();
    event.stopImmediatePropagation();

    const form = button.closest("#commessa-mobile-plant-form");
    if (!form) return;
    const status = form.querySelector("#commessa-mobile-geocode-status");
    const setStatus = (message) => { if (status) status.textContent = message; };
    restoreMobilePlantEditor(form);

    if (!navigator.geolocation) {
      setStatus("La posizione automatica non è disponibile su questo dispositivo.");
      return;
    }

    let requestFinished = false;
    const keepEditorOpen = () => {
      if (!requestFinished && !document.hidden) restoreMobilePlantEditor(form);
    };
    const cleanup = () => {
      requestFinished = true;
      window.removeEventListener("focus", keepEditorOpen);
      document.removeEventListener("visibilitychange", keepEditorOpen);
    };
    window.addEventListener("focus", keepEditorOpen);
    document.addEventListener("visibilitychange", keepEditorOpen);

    button.disabled = true;
    button.textContent = "Rilevamento posizione…";
    setStatus("Rilevamento GPS in corso…");

    navigator.geolocation.getCurrentPosition(position => {
      cleanup();
      if (!document.contains(button) || button.closest("#commessa-mobile-plant-form") !== form) return;
      restoreMobilePlantEditor(form);
      const latitude = Number(position?.coords?.latitude);
      const longitude = Number(position?.coords?.longitude);
      if (!Number.isFinite(latitude) || !Number.isFinite(longitude) || latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) {
        button.disabled = false;
        button.textContent = "📍 Usa la mia posizione";
        setStatus("Il telefono ha restituito coordinate non valide. Riprova.");
        return;
      }

      const latitudeInput = form.querySelector('[data-v2-field="latitudine"]');
      const longitudeInput = form.querySelector('[data-v2-field="longitudine"]');
      if (latitudeInput) latitudeInput.value = latitude.toFixed(6);
      if (longitudeInput) longitudeInput.value = longitude.toFixed(6);
      form.dataset.mobileGeocodeKey = "";
      latitudeInput?.dispatchEvent(new Event("input", { bubbles: true }));
      longitudeInput?.dispatchEvent(new Event("input", { bubbles: true }));

      restoreMobilePlantEditor(form);
      button.disabled = false;
      button.textContent = "✓ Posizione acquisita";
      setStatus("Posizione acquisita. Ricerca automatica di Comune e Via…");
    }, error => {
      cleanup();
      if (!document.contains(button) || button.closest("#commessa-mobile-plant-form") !== form) return;
      restoreMobilePlantEditor(form);
      button.disabled = false;
      button.textContent = "📍 Usa la mia posizione";
      if (error?.code === 1) setStatus("Non è stato possibile accedere alla posizione. Consenti la localizzazione, verifica che il GPS sia attivo e riprova.");
      else if (error?.code === 2) setStatus("Posizione non disponibile. Verifica che il GPS del telefono sia attivo e riprova.");
      else if (error?.code === 3) setStatus("Rilevamento della posizione scaduto. Spostati in un punto con più segnale e riprova.");
      else setStatus("Non è stato possibile rilevare la posizione. Verifica il GPS e riprova.");
    }, { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 });
  }

  function updateAvailability() {
    const available = typeof canManageData === "function"
      && canManageData()
      && typeof selectedCommessaId !== "undefined"
      && Boolean(String(selectedCommessaId || "").trim());
    wrap.classList.toggle("hidden", !available);
  }

  async function openCommessaManagement() {
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

  ensureMobileView();
  document.addEventListener("click", handleMobileCurrentLocation, true);
  document.getElementById("commessa-plants-menu")?.remove();
  toggle.removeAttribute("aria-haspopup");
  toggle.removeAttribute("aria-expanded");
  toggle.removeAttribute("aria-controls");
  toggle.setAttribute("aria-label", "Apri gestione commessa");
  toggle.setAttribute("title", "Gestione commessa");
  toggle.addEventListener("click", (event) => {
    event.preventDefault();
    void openCommessaManagement();
  });

  new MutationObserver(updateAvailability).observe(title, { childList: true, characterData: true, subtree: true });
  window.addEventListener("pageshow", updateAvailability);
  updateAvailability();
})();

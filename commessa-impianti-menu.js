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
  let opening = false;
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

  function handleMobileCurrentLocation(event) {
    const button = event.target.closest?.("#commessa-mobile-current-location");
    if (!button) return;

    // Questo pulsante vive dentro il form di inserimento impianto. Su iOS il
    // click poteva proseguire verso listener globali e riportare alla home della
    // gestione commessa. Lo intercettiamo qui, prima del bubbling, senza toccare
    // FATTO, WhatsApp o il modello dati degli impianti.
    event.preventDefault();
    event.stopImmediatePropagation();

    const form = button.closest("#commessa-mobile-plant-form");
    if (!form) return;
    const status = form.querySelector("#commessa-mobile-geocode-status");
    const setStatus = (message) => { if (status) status.textContent = message; };

    if (!navigator.geolocation) {
      setStatus("La posizione automatica non è disponibile su questo dispositivo.");
      return;
    }

    button.disabled = true;
    button.textContent = "Rilevamento posizione…";
    setStatus("Rilevamento GPS in corso…");

    navigator.geolocation.getCurrentPosition(position => {
      if (!document.contains(button) || button.closest("#commessa-mobile-plant-form") !== form) return;
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

      // Riutilizza il flusso già esistente di accounting-v2: gli eventi input
      // avviano la compilazione automatica di Comune e Via senza cambiare vista.
      latitudeInput?.dispatchEvent(new Event("input", { bubbles: true }));
      longitudeInput?.dispatchEvent(new Event("input", { bubbles: true }));

      button.disabled = false;
      button.textContent = "✓ Posizione acquisita";
      setStatus("Posizione acquisita. Ricerca automatica di Comune e Via…");
    }, error => {
      if (!document.contains(button) || button.closest("#commessa-mobile-plant-form") !== form) return;
      button.disabled = false;
      button.textContent = "📍 Usa la mia posizione";
      if (error?.code === 1) setStatus("Non è stato possibile accedere alla posizione. Consenti la localizzazione, verifica che il GPS sia attivo e riprova.");
      else if (error?.code === 2) setStatus("Posizione non disponibile. Verifica che il GPS del telefono sia attivo e riprova.");
      else if (error?.code === 3) setStatus("Rilevamento della posizione scaduto. Spostati in un punto con più segnale e riprova.");
      else setStatus("Non è stato possibile rilevare la posizione. Verifica il GPS e riprova.");
    }, { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 });
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

  function showMobileOpeningState(commessa) {
    const mobile = document.getElementById("commessa-mobile-management");
    if (!mobile) return;
    mobile.classList.remove("hidden");
    mobile.setAttribute("aria-hidden", "false");
    mobile.setAttribute("aria-busy", "true");
    document.getElementById("commessa-mobile-management-home")?.classList.remove("hidden");
    document.getElementById("commessa-mobile-plant-list")?.classList.add("hidden");
    document.getElementById("commessa-mobile-plant-editor")?.classList.add("hidden");
    const title = document.getElementById("commessa-mobile-management-title");
    const meta = document.getElementById("commessa-mobile-management-meta");
    const stats = document.getElementById("commessa-mobile-management-stats");
    if (title) title.textContent = commessa?.nome || "Gestione commessa";
    if (meta) meta.textContent = `Cod. ${commessa?.codice || "—"}`;
    if (stats) stats.innerHTML = '<span><b>…</b> caricamento dati commessa</span>';
    mobile.querySelectorAll('[data-commessa-mobile-action]').forEach(button => { button.disabled = true; });
  }

  function finishMobileOpeningState() {
    const mobile = document.getElementById("commessa-mobile-management");
    if (!mobile) return;
    mobile.removeAttribute("aria-busy");
    mobile.querySelectorAll('[data-commessa-mobile-action]').forEach(button => { button.disabled = false; });
  }

  function showMobileOpeningError(error) {
    const mobile = document.getElementById("commessa-mobile-management");
    const stats = document.getElementById("commessa-mobile-management-stats");
    if (mobile) {
      mobile.classList.remove("hidden");
      mobile.setAttribute("aria-hidden", "false");
      mobile.removeAttribute("aria-busy");
    }
    if (stats) stats.innerHTML = `<span><b>⚠️</b> ${String(error?.message || "Caricamento non riuscito. Chiudi e riprova.").replace(/[&<>"']/g, char => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]))}</span>`;
  }

  async function openCommessaManagement(toggle) {
    if (opening) return;
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
    opening = true;
    toggle.disabled = true;
    try {
      ensureMobileView();
      openManagementPanel("commesse");
      showMobileOpeningState(commessa);
      const opened = await window.AccountingV2.openMobileHub(commessa);
      if (!opened) throw new Error("Il caricamento della commessa è stato interrotto. Riprova.");
      finishMobileOpeningState();
    } catch (error) {
      console.error("Apertura gestione commessa non riuscita:", error);
      showMobileOpeningError(error);
    } finally {
      opening = false;
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
    document.addEventListener("click", handleMobileCurrentLocation, true);
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

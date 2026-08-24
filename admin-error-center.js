(() => {
  "use strict";

  if (window.HeraAdminErrorCenter?.installed) return;

  const VERSION = "1.0.0";
  const REGION = "europe-west1";
  const ADMIN_EMAIL = "ionut29019@gmail.com";
  const FUNCTIONS = Object.freeze({
    summary: "getErrorCenterSummary",
    dashboard: "getErrorCenterDashboard",
    seen: "markErrorCenterSeen",
    update: "updateErrorCenterStatus"
  });

  const state = {
    user: null,
    admin: false,
    items: [],
    counts: {},
    selectedId: "",
    pendingTarget: "",
    menuAttempts: 0,
    authAttempts: 0,
    loading: false
  };

  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[character]);

  function isAdminUser(user = state.user) {
    try {
      if (typeof canManageData === "function" && canManageData()) return true;
    } catch (_) {}
    return String(user?.email || "").trim().toLowerCase() === ADMIN_EMAIL;
  }

  function functionsCallable(name) {
    if (!window.firebase?.apps?.length || !window.firebase?.functions) return null;
    try { return window.firebase.app().functions(REGION).httpsCallable(name); } catch (_) { return null; }
  }

  async function callFunction(name, data = {}) {
    const invoke = functionsCallable(name);
    if (!invoke) throw new Error("Servizi Firebase non ancora disponibili. Riprova tra poco.");
    const response = await invoke(data);
    return response?.data || {};
  }

  function ensureStyle() {
    if (document.querySelector('link[data-admin-error-center-style]')) return;
    const link = document.createElement("link");
    link.rel = "stylesheet";
    link.href = "./admin-error-center.css?v=20260824a";
    link.dataset.adminErrorCenterStyle = "1";
    document.head.appendChild(link);
  }

  function closeSideMenu() {
    const menu = document.getElementById("side-menu");
    const overlay = document.getElementById("menu-overlay");
    menu?.classList.add("hidden");
    menu?.setAttribute("aria-hidden", "true");
    overlay?.classList.add("hidden");
  }

  function toolsSection() {
    const title = document.getElementById("menu-strumenti-title");
    return title?.closest?.(".menu-section") || title?.parentElement || null;
  }

  function createMenuButton(id, icon, label) {
    const button = document.createElement("button");
    button.id = id;
    button.type = "button";
    button.className = "btn menu-title-btn";
    button.innerHTML = `<span class="menu-item-icon" aria-hidden="true">${icon}</span><span>${esc(label)}</span>`;
    return button;
  }

  function ensureMenuButtons() {
    const section = toolsSection();
    if (!section) {
      if (state.menuAttempts < 12) {
        state.menuAttempts += 1;
        window.setTimeout(ensureMenuButtons, 500);
      }
      return;
    }

    state.menuAttempts = 0;
    let reportButton = document.getElementById("open-app-bug-report-btn");
    if (!reportButton) {
      reportButton = createMenuButton("open-app-bug-report-btn", "🐞", "Segnala problema app");
      reportButton.addEventListener("click", openBugReport);
      section.appendChild(reportButton);
    }
    reportButton.hidden = !state.user;

    let adminButton = document.getElementById("open-admin-error-center-btn");
    if (state.admin && !adminButton) {
      adminButton = createMenuButton("open-admin-error-center-btn", "⚠️", "Centro errori");
      adminButton.classList.add("hera-error-menu-button");
      const badge = document.createElement("span");
      badge.className = "hera-error-menu-badge";
      badge.dataset.errorCenterBadge = "1";
      badge.hidden = true;
      adminButton.appendChild(badge);
      adminButton.addEventListener("click", () => void openCenter());
      section.appendChild(adminButton);
    }
    if (adminButton) adminButton.hidden = !state.admin;
  }

  function ensureDialogs() {
    if (!document.getElementById("hera-error-center-dialog")) {
      const dialog = document.createElement("dialog");
      dialog.id = "hera-error-center-dialog";
      dialog.className = "hera-error-dialog";
      dialog.innerHTML = `
        <div class="hera-error-shell">
          <header class="hera-error-head">
            <div><h2>⚠️ Centro errori</h2><p>Errori, rallentamenti, blocchi e segnalazioni rilevati nell’app.</p></div>
            <div class="hera-error-head-actions"><button class="btn" type="button" data-error-refresh>AGGIORNA</button><button class="btn" type="button" data-error-close>CHIUDI</button></div>
          </header>
          <section class="hera-error-summary" data-error-summary></section>
          <section class="hera-error-toolbar">
            <input type="search" data-error-query placeholder="Cerca funzione, pagina, messaggio o commessa" autocomplete="off">
            <select data-error-status aria-label="Filtra per stato"><option value="all">Tutti gli stati</option><option value="open">Aperti</option><option value="in_verification">In verifica</option><option value="resolved">Risolti</option><option value="ignored">Ignorati</option></select>
            <select data-error-severity aria-label="Filtra per gravità"><option value="all">Tutte le gravità</option><option value="critical">Critici</option><option value="high">Alti</option><option value="medium">Medi</option><option value="low">Bassi</option><option value="info">Informativi</option></select>
            <button class="btn btn-primary" type="button" data-error-apply>APPLICA</button>
          </section>
          <section class="hera-error-body">
            <div class="hera-error-list" data-error-list><div class="hera-error-loading">Caricamento errori…</div></div>
            <div class="hera-error-detail" data-error-detail><div class="hera-error-empty">Seleziona un errore per vedere tutti i dettagli.</div></div>
          </section>
        </div>`;
      document.body.appendChild(dialog);
      dialog.addEventListener("cancel", (event) => { event.preventDefault(); closeCenter(); });
      dialog.addEventListener("click", handleCenterClick);
    }

    if (!document.getElementById("hera-bug-report-dialog")) {
      const dialog = document.createElement("dialog");
      dialog.id = "hera-bug-report-dialog";
      dialog.className = "hera-bug-dialog";
      dialog.innerHTML = `
        <div class="hera-bug-shell">
          <header class="hera-bug-head"><div><h2>🐞 Segnala problema app</h2><p>La segnalazione arriverà al Centro errori dell’amministratore.</p></div><button class="btn" type="button" data-bug-close>CHIUDI</button></header>
          <form class="hera-bug-form" data-bug-form>
            <label>Titolo del problema<input name="title" maxlength="180" required placeholder="Es. Il pulsante Home non reagisce"></label>
            <label>Gravità<select name="severity"><option value="medium">Media</option><option value="low">Bassa</option><option value="high">Alta</option><option value="critical">Critica: impedisce di lavorare</option></select></label>
            <label>Cosa è successo<textarea name="description" maxlength="3000" required placeholder="Descrivi il problema senza inserire password, PIN o dati personali."></textarea></label>
            <label>Passaggi per riprodurlo<textarea name="steps" maxlength="2200" placeholder="Es. Apro la commessa, premo Impianti consigliati, poi Home…"></textarea></label>
            <label>Cosa doveva succedere<textarea name="expected" maxlength="1800" placeholder="Descrivi il comportamento corretto atteso."></textarea></label>
            <p class="hera-bug-feedback" data-bug-feedback role="status" aria-live="polite"></p>
            <div class="hera-bug-actions"><button class="btn" type="button" data-bug-close>ANNULLA</button><button class="btn btn-primary" type="submit" data-bug-submit>INVIA SEGNALAZIONE</button></div>
          </form>
        </div>`;
      document.body.appendChild(dialog);
      dialog.addEventListener("cancel", (event) => { event.preventDefault(); closeBugReport(); });
      dialog.addEventListener("click", (event) => {
        if (event.target.closest?.("[data-bug-close]")) closeBugReport();
      });
      dialog.querySelector("[data-bug-form]")?.addEventListener("submit", submitBugReport);
    }
  }

  function showDialog(dialog) {
    if (!dialog) return;
    if (typeof dialog.showModal === "function") {
      if (!dialog.open) dialog.showModal();
    } else {
      dialog.setAttribute("open", "");
    }
  }

  function hideDialog(dialog) {
    if (!dialog) return;
    if (typeof dialog.close === "function" && dialog.open) dialog.close();
    else dialog.removeAttribute("open");
  }

  function setBadge(count) {
    const badge = document.querySelector("[data-error-center-badge]");
    if (!badge) return;
    const numeric = Math.max(0, Number(count || 0));
    badge.hidden = numeric === 0;
    badge.textContent = numeric > 99 ? "99+" : String(numeric);
    document.getElementById("open-admin-error-center-btn")?.setAttribute(
      "aria-label",
      `Apri Centro errori${numeric ? `, ${numeric} nuovi avvisi` : ""}`
    );
  }

  async function refreshSummary() {
    if (!state.admin || !state.user) return;
    try {
      const summary = await callFunction(FUNCTIONS.summary);
      setBadge(summary.unseenAlerts || 0);
    } catch (error) {
      console.warn("Riepilogo Centro errori non disponibile:", error);
    }
  }

  const severityLabel = (severity) => ({ critical: "Critico", high: "Alto", medium: "Medio", low: "Basso", info: "Info" })[severity] || "Medio";
  const statusLabel = (status) => ({ open: "Aperto", in_verification: "In verifica", resolved: "Risolto", ignored: "Ignorato" })[status] || "Aperto";

  function formatDate(value) {
    const date = new Date(value || 0);
    if (!Number.isFinite(date.getTime())) return "—";
    return date.toLocaleString("it-IT", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
  }

  function renderSummary() {
    const root = document.querySelector("[data-error-summary]");
    if (!root) return;
    const counts = state.counts || {};
    root.innerHTML = [
      ["Aperti", counts.open || 0, ""],
      ["In verifica", counts.inVerification || 0, ""],
      ["Critici", counts.critical || 0, "is-critical"],
      ["Alti", counts.high || 0, "is-high"],
      ["Risolti", counts.resolved || 0, ""],
      ["Ignorati", counts.ignored || 0, ""]
    ].map(([label, value, className]) => `<div class="hera-error-summary-card ${className}"><strong>${Number(value)}</strong><span>${esc(label)}</span></div>`).join("");
  }

  function renderList() {
    const root = document.querySelector("[data-error-list]");
    if (!root) return;
    if (state.loading) {
      root.innerHTML = '<div class="hera-error-loading">Caricamento errori…</div>';
      return;
    }
    if (!state.items.length) {
      root.innerHTML = '<div class="hera-error-empty"><strong>Nessun errore trovato.</strong><br>Modifica i filtri oppure aggiorna il pannello.</div>';
      return;
    }
    root.innerHTML = state.items.map((item) => `
      <button type="button" class="hera-error-row ${state.selectedId === item.id ? "is-selected" : ""}" data-error-id="${esc(item.id)}">
        <span class="hera-error-dot" data-severity="${esc(item.severity)}"></span>
        <span class="hera-error-row-copy"><strong>${esc(item.title || item.category || "Errore app")}</strong><span>${esc(item.lastMessage || "Nessun messaggio")}</span><small>${esc(statusLabel(item.status))} · ${esc(formatDate(item.lastSeenAt))} · ${esc(item.lastPlatform || "dispositivo non noto")}</small></span>
        <span class="hera-error-count">${Number(item.occurrences || 0)}×</span>
      </button>`).join("");
  }

  function eventDetails(event, index) {
    const technical = {
      source: event.source || "",
      line: event.line || null,
      column: event.column || null,
      metadata: event.metadata || {},
      breadcrumbs: event.breadcrumbs || []
    };
    return `<details class="hera-error-event" ${index === 0 ? "open" : ""}><summary>${esc(formatDate(event.occurredAt))} · ${esc(event.platform || "dispositivo")} · ${esc(event.userName || event.userEmailMasked || "utente")}</summary><p>${esc(event.message || "")}</p><pre>${esc(JSON.stringify(technical, null, 2))}</pre></details>`;
  }

  function selectedItem() {
    return state.items.find((item) => item.id === state.selectedId) || null;
  }

  function renderDetail() {
    const root = document.querySelector("[data-error-detail]");
    if (!root) return;
    const item = selectedItem();
    if (!item) {
      root.innerHTML = '<div class="hera-error-empty">Seleziona un errore per vedere tutti i dettagli.</div>';
      return;
    }
    root.innerHTML = `
      <span class="hera-error-kicker">${esc(item.category || "ERRORE APP")}</span>
      <h3>${esc(item.title || "Errore app")}</h3>
      <div class="hera-error-tags"><span class="hera-error-tag" data-severity="${esc(item.severity)}">${esc(severityLabel(item.severity))}</span><span class="hera-error-tag">${esc(statusLabel(item.status))}</span><span class="hera-error-tag">${Number(item.occurrences || 0)} occorrenze</span><span class="hera-error-tag">${Number(item.affectedUsers || 0)} utenti</span></div>
      <dl class="hera-error-detail-grid">
        <div><dt>Funzione</dt><dd>${esc(item.feature || "—")}</dd></div><div><dt>Ultima schermata</dt><dd>${esc(item.lastActiveView || item.lastPage || "—")}</dd></div>
        <div><dt>Prima comparsa</dt><dd>${esc(formatDate(item.firstSeenAt))}</dd></div><div><dt>Ultima comparsa</dt><dd>${esc(formatDate(item.lastSeenAt))}</dd></div>
        <div><dt>Dispositivo</dt><dd>${esc(item.lastPlatform || "—")}</dd></div><div><dt>Versione app</dt><dd>${esc(item.lastAppVersion || "—")}</dd></div>
        <div><dt>Utente</dt><dd>${esc(item.lastUserName || item.lastUserEmailMasked || "—")}</dd></div><div><dt>Durata / tocchi</dt><dd>${item.lastDurationMs ? `${Number(item.lastDurationMs)} ms` : "—"}${item.lastTapCount ? ` · ${Number(item.lastTapCount)} tocchi` : ""}</dd></div>
        <div><dt>Commessa</dt><dd>${esc(item.commessaName || item.commessaId || "—")}</dd></div><div><dt>Impianto</dt><dd>${esc(item.impiantoId || "—")}</dd></div>
      </dl>
      <h4>Problema</h4><div class="hera-error-message">${esc(item.lastMessage || "Nessun messaggio")}</div>
      <h4>Azione tecnica consigliata</h4><div class="hera-error-action">${esc(item.diagnosisAction || "Analizzare il dettaglio tecnico.")}</div>
      <h4>Stack tecnico</h4><pre class="hera-error-stack">${esc(item.lastStack || "Non disponibile")}</pre>
      <h4>Ultimi eventi registrati</h4><div class="hera-error-events">${(item.recentEvents || []).map(eventDetails).join("") || '<p class="muted">Nessun evento dettagliato disponibile.</p>'}</div>
      <section class="hera-error-admin">
        <label>Stato<select data-error-detail-status>${["open", "in_verification", "resolved", "ignored"].map((status) => `<option value="${status}" ${item.status === status ? "selected" : ""}>${esc(statusLabel(status))}</option>`).join("")}</select></label>
        <label>Nota amministratore<textarea class="hera-error-note" data-error-note rows="4" maxlength="3000" placeholder="Annota verifica, causa o correzione applicata.">${esc(item.adminNote || "")}</textarea></label>
        <p class="hera-error-status" data-error-feedback></p>
        <div class="hera-error-admin-actions"><button class="btn btn-primary" type="button" data-error-save>SALVA STATO</button><button class="btn" type="button" data-error-copy>COPIA DIAGNOSTICA</button></div>
      </section>`;
  }

  async function loadDashboard() {
    if (!state.admin || state.loading) return;
    state.loading = true;
    renderList();
    const query = document.querySelector("[data-error-query]")?.value || "";
    const status = document.querySelector("[data-error-status]")?.value || "all";
    const severity = document.querySelector("[data-error-severity]")?.value || "all";
    try {
      const result = await callFunction(FUNCTIONS.dashboard, { query, status, severity, limit: 120 });
      state.items = Array.isArray(result.items) ? result.items : [];
      state.counts = result.counts || {};
      if (state.pendingTarget && state.items.some((item) => item.id === state.pendingTarget)) state.selectedId = state.pendingTarget;
      else if (!state.items.some((item) => item.id === state.selectedId)) state.selectedId = state.items[0]?.id || "";
      state.pendingTarget = "";
    } catch (error) {
      console.error("Caricamento Centro errori fallito:", error);
      state.items = [];
      state.counts = {};
      const root = document.querySelector("[data-error-list]");
      if (root) root.innerHTML = `<div class="hera-error-empty"><strong>Centro errori non disponibile.</strong><br>${esc(error?.message || "Controlla il deploy delle Cloud Functions.")}</div>`;
    } finally {
      state.loading = false;
      renderSummary();
      renderList();
      renderDetail();
    }
  }

  async function markSeen() {
    setBadge(0);
    try { await callFunction(FUNCTIONS.seen); } catch (error) { console.warn("Conferma lettura Centro errori non sincronizzata:", error); }
  }

  async function openCenter(groupId = "") {
    if (!state.admin) {
      window.alert("Il Centro errori è riservato all’amministratore.");
      return;
    }
    ensureStyle();
    ensureDialogs();
    closeSideMenu();
    if (groupId) state.pendingTarget = String(groupId);
    showDialog(document.getElementById("hera-error-center-dialog"));
    void markSeen();
    await loadDashboard();
  }

  function closeCenter() {
    hideDialog(document.getElementById("hera-error-center-dialog"));
  }

  function openBugReport() {
    if (!state.user) {
      window.alert("Accedi all’app prima di inviare una segnalazione.");
      return;
    }
    ensureStyle();
    ensureDialogs();
    closeSideMenu();
    const dialog = document.getElementById("hera-bug-report-dialog");
    const form = dialog?.querySelector("[data-bug-form]");
    form?.reset();
    const feedback = dialog?.querySelector("[data-bug-feedback]");
    if (feedback) feedback.textContent = "";
    showDialog(dialog);
  }

  function closeBugReport() {
    hideDialog(document.getElementById("hera-bug-report-dialog"));
  }

  async function submitBugReport(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const submit = form.querySelector("[data-bug-submit]");
    const feedback = form.querySelector("[data-bug-feedback]");
    if (submit?.disabled) return;
    const data = new FormData(form);
    const payload = {
      title: data.get("title"),
      severity: data.get("severity"),
      description: data.get("description"),
      steps: data.get("steps"),
      expected: data.get("expected")
    };
    if (!String(payload.title || "").trim() || !String(payload.description || "").trim()) {
      if (feedback) feedback.textContent = "Titolo e descrizione sono obbligatori.";
      return;
    }
    if (submit) submit.disabled = true;
    if (feedback) feedback.textContent = "Invio segnalazione…";
    try {
      const monitor = window.HeraAppErrorMonitor;
      if (!monitor?.reportManual) throw new Error("Monitor errori non ancora caricato. Riapri l’app e riprova.");
      const result = await monitor.reportManual(payload);
      if (feedback) feedback.textContent = result.sent
        ? "✅ Segnalazione inviata all’amministratore."
        : "✅ Segnalazione salvata: verrà inviata automaticamente appena possibile.";
      form.reset();
    } catch (error) {
      if (feedback) feedback.textContent = `⚠️ ${error?.message || "Invio non riuscito."}`;
    } finally {
      if (submit) submit.disabled = false;
    }
  }

  async function saveSelectedStatus() {
    const item = selectedItem();
    if (!item) return;
    const status = document.querySelector("[data-error-detail-status]")?.value || item.status;
    const adminNote = document.querySelector("[data-error-note]")?.value || "";
    const feedback = document.querySelector("[data-error-feedback]");
    if (feedback) feedback.textContent = "Salvataggio…";
    try {
      await callFunction(FUNCTIONS.update, { groupId: item.id, status, adminNote });
      if (feedback) feedback.textContent = "✅ Stato aggiornato.";
      item.status = status;
      item.adminNote = adminNote;
      await loadDashboard();
    } catch (error) {
      if (feedback) feedback.textContent = `⚠️ ${error?.message || "Salvataggio non riuscito."}`;
    }
  }

  async function copySelectedDiagnostic() {
    const item = selectedItem();
    if (!item) return;
    const text = JSON.stringify(item, null, 2);
    try {
      if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(text);
      else {
        const textarea = document.createElement("textarea");
        textarea.value = text;
        textarea.style.position = "fixed";
        textarea.style.opacity = "0";
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand("copy");
        textarea.remove();
      }
      const feedback = document.querySelector("[data-error-feedback]");
      if (feedback) feedback.textContent = "✅ Diagnostica copiata.";
    } catch (error) {
      window.alert(`Copia non riuscita: ${error?.message || "errore sconosciuto"}`);
    }
  }

  function handleCenterClick(event) {
    if (event.target.closest?.("[data-error-close]")) return closeCenter();
    if (event.target.closest?.("[data-error-refresh], [data-error-apply]")) return void loadDashboard();
    const row = event.target.closest?.("[data-error-id]");
    if (row) {
      state.selectedId = row.dataset.errorId || "";
      renderList();
      renderDetail();
      return;
    }
    if (event.target.closest?.("[data-error-save]")) return void saveSelectedStatus();
    if (event.target.closest?.("[data-error-copy]")) return void copySelectedDiagnostic();
  }

  function installForUser(user) {
    state.user = user || null;
    state.admin = Boolean(user && isAdminUser(user));
    ensureStyle();
    ensureDialogs();
    ensureMenuButtons();
    if (state.admin) window.setTimeout(() => void refreshSummary(), 900);
    else setBadge(0);

    if (state.admin) {
      const params = new URLSearchParams(location.search || "");
      const groupId = params.get("openErrorCenter");
      if (groupId) {
        params.delete("openErrorCenter");
        const nextUrl = `${location.pathname}${params.toString() ? `?${params}` : ""}${location.hash || ""}`;
        try { history.replaceState(history.state, "", nextUrl); } catch (_) {}
        window.setTimeout(() => void openCenter(groupId), 0);
      }
    }
  }

  function bindAuth() {
    try {
      const auth = window.firebase?.auth?.();
      if (!auth) throw new Error("Firebase Auth non pronto");
      auth.onAuthStateChanged(installForUser);
      return;
    } catch (_) {}
    if (state.authAttempts < 20) {
      state.authAttempts += 1;
      window.setTimeout(bindAuth, 500);
    }
  }

  function install() {
    ensureStyle();
    ensureDialogs();
    ensureMenuButtons();
    bindAuth();
  }

  window.addEventListener("hera:open-notification-destination", (event) => {
    const destination = String(event.detail?.destination || "").toUpperCase();
    if (destination === "ERROR_CENTER") void openCenter(event.detail?.target || "");
  });

  try {
    navigator.serviceWorker?.addEventListener?.("message", (event) => {
      const notification = event.data?.notification || event.data?.message || {};
      if (String(notification.destination || notification.rawData?.destination || "").toLowerCase() === "error-center") {
        void openCenter(notification.rawData?.groupId || notification.groupId || "");
      }
    });
  } catch (_) {}

  window.HeraAdminErrorCenter = Object.freeze({
    installed: true,
    version: VERSION,
    open: openCenter,
    openBugReport,
    refresh: loadDashboard,
    refreshSummary
  });

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", install, { once: true });
  else install();
})();

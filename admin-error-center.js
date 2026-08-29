(() => {
  "use strict";

  if (window.HeraAdminErrorCenter?.installed) return;

  const VERSION = "1.5.0";
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
    loading: false,
    dataVerified: false,
    loadError: "",
    serverTime: "",
    truncated: false,
    impactFilter: "all"
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
    link.href = "./admin-error-center.css?v=20260824b";
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
          <section class="hera-error-health" data-error-health role="status" aria-live="polite"></section>
          <section class="hera-error-summary" data-error-summary></section>
          <section class="hera-error-impact-summary" data-error-impact-summary aria-label="Categorie di impatto degli errori"></section>
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

  const IMPACT_CATEGORIES = Object.freeze([
    { id: "blocking", icon: "⛔", label: "Blocca l’app", note: "Crash, schermate bloccate o comandi che non rispondono." },
    { id: "commesse", icon: "📂", label: "Rischio commesse", note: "Può mescolare, perdere o assegnare male dati di commesse e impianti." },
    { id: "performance", icon: "🐢", label: "Rallenta l’app", note: "Listener, render, timeout o attività ripetute che degradano le prestazioni." },
    { id: "sync", icon: "💾", label: "Salvataggio / sync", note: "Errori di Firestore, rete, offline o scrittura dei dati." },
    { id: "auth", icon: "🔐", label: "Login / permessi", note: "Accesso, sessione, autorizzazioni o credenziali." },
    { id: "notifications", icon: "🔔", label: "Notifiche", note: "Push, FCM, VAPID o configurazione delle notifiche." },
    { id: "map", icon: "🗺️", label: "Mappa / GPS", note: "Coordinate, geolocalizzazione, Leaflet o navigazione." },
    { id: "other", icon: "⚠️", label: "Altri errori", note: "Errore non riconducibile con sufficiente sicurezza alle altre categorie." }
  ]);

  function normalizeImpactText(item) {
    const events = Array.isArray(item?.recentEvents) ? item.recentEvents : [];
    return [
      item?.title, item?.category, item?.lastMessage, item?.feature, item?.lastActiveView,
      item?.lastPage, item?.diagnosisAction, item?.lastStack,
      ...events.flatMap((event) => [event?.message, event?.source])
    ].filter(Boolean).join(" ").toLowerCase();
  }

  function impactCategory(item) {
    const text = normalizeImpactText(item);
    const match = (pattern) => pattern.test(text);

    if (match(/\b(crash|fatal|freeze|frozen|hang|hung|unresponsive|uncaught)\b|blocca(?:to|ta|re)?|bloccato|schermata bianca|non risponde|non reagisce|maximum call stack|out of memory/)) return IMPACT_CATEGORIES[0];
    if (match(/commessa|impianto/) && match(/mescol|sovrascr|perdit|pers[oi]|cancell|elimin|spostat|assegnat.*sbagli|id errat|integrit|isolament|altra commessa|commessa sbagliata|impianti di un altra|impianti di un'altra/)) return IMPACT_CATEGORIES[1];
    if (match(/notific|notification|push|messaging|fcm|vapid|applicationserverkey|p-256|service.?worker.*push/)) return IMPACT_CATEGORIES[5];
    if (match(/mappa|\bmap\b|gps|geoloc|leaflet|google maps|coordinate|latitud|longitud|navigaz/)) return IMPACT_CATEGORIES[6];
    if (match(/\bauth\b|login|permission-denied|permess|unauthor|forbidden|credential|sessione|token.*scad|accesso negato/)) return IMPACT_CATEGORIES[4];
    if (match(/rallent|\bslow\b|performance|long task|listener duplic|listener.*multip|render.*ripet|memory leak|troppe letture|richieste ripet|latency|latenza|timeout/)) return IMPACT_CATEGORIES[2];
    if (match(/firestore|database|storage|sincron|\bsync\b|offline|salvat|scritt|\bwrite\b|network|rete|quota/)) return IMPACT_CATEGORIES[3];
    return IMPACT_CATEGORIES[7];
  }

  function impactItems() {
    return state.impactFilter === "all" ? state.items : state.items.filter((item) => impactCategory(item).id === state.impactFilter);
  }

  function chatGptCategoryItems() {
    const selected = selectedItem();
    const impactId = state.impactFilter !== "all" ? state.impactFilter : impactCategory(selected).id;
    return state.items.filter((item) => impactCategory(item).id === impactId);
  }

  function uniqueRecentEvents(events = []) {
    const seen = new Set();
    return events.filter((event) => {
      const key = [event?.message, event?.source, event?.line, event?.platform, event?.userEmailMasked].map((value) => String(value || "")).join("|");
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function chatGptCategoryText() {
    const items = chatGptCategoryItems();
    const category = items[0] ? impactCategory(items[0]) : IMPACT_CATEGORIES.find((item) => item.id === state.impactFilter) || IMPACT_CATEGORIES[7];
    const payload = items.map((item) => ({
      ...item,
      recentEvents: uniqueRecentEvents(item.recentEvents || [])
    }));
    const totalOccurrences = items.reduce((sum, item) => sum + Math.max(1, Number(item.occurrences || 0)), 0);
    return [
      "Analizza questo blocco completo del Centro errori della mia app Varga Cantieri.",
      "Categoria di impatto: " + category.icon + " " + category.label + ".",
      "Gruppi inclusi: " + items.length + ". Occorrenze complessive dichiarate: " + totalOccurrences + ".",
      "Distingui gli errori realmente diversi dai duplicati, confronta firstSeenAt/lastSeenAt con le versioni dei file, indica quali sono storici o già risolti e quali richiedono ancora una correzione.",
      "Per ogni problema ancora attivo proponi la modifica minima e sicura. Non modificare FATTO, WhatsApp/WHAZZUP, dati di commesse, impianti, personale, utenti, squadre, calendario o ore se non sono direttamente coinvolti.",
      "Non considerare automaticamente feature o lastActiveView come causa: verifica sempre message, stack e source.",
      "Rispondi in italiano con: causa, stato (storico/duplicato/attivo), priorità e azione consigliata.",
      "",
      "DATI DELLA CATEGORIA:",
      JSON.stringify(payload, null, 2)
    ].join("\n");
  }

  function repairAdvice(item) {
    const text = normalizeImpactText(item);
    const impact = impactCategory(item).id;
    const generic = {
      title: "Analisi guidata",
      summary: item?.diagnosisAction || "Analizzare il dettaglio tecnico e riprodurre l'errore in modo controllato.",
      steps: [
        "Riprodurre l'errore una sola volta e verificare l'ultimo evento registrato.",
        "Controllare stack, pagina, funzione e piattaforma coinvolta.",
        "Applicare una correzione su branch dedicato e verificare i flussi critici prima del merge."
      ],
      safeAutoRepair: false,
      repairNote: "Richiede una modifica tecnica verificata: il Centro errori non cambia automaticamente il codice sorgente."
    };

    if (/applicationserverkey|p-256|vapid/.test(text)) return {
      title: "Configurazione Web Push / VAPID",
      summary: "La chiave applicationServerKey non è una chiave pubblica P-256 valida. Se le notifiche push non servono, la soluzione più sicura è impedire la registrazione automatica del Web Push; altrimenti va configurata una chiave VAPID pubblica valida.",
      steps: [
        "Bloccare la subscribe Push quando Web Push è disattivato o non configurato.",
        "Non richiedere il permesso notifiche automaticamente all'avvio.",
        "Se si vuole usare Web Push, validare la VAPID public key prima di chiamare pushManager.subscribe().",
        "Pulire eventuali subscription locali non valide e verificare su PWA iOS."
      ],
      safeAutoRepair: true,
      repairNote: "Riparazione sicura disponibile sul dispositivo: rimuove una subscription Web Push locale non valida e disattiva i tentativi automatici per questa installazione. Non modifica FATTO, WHAZZUP o i dati delle commesse."
    };

    if (impact === "performance") return {
      title: "Riduzione rallentamenti",
      summary: "Probabile lavoro ripetuto, listener duplicato, timeout o render eccessivo.",
      steps: [
        "Individuare listener/observer registrati più volte e garantire una sola sottoscrizione attiva.",
        "Evitare refresh completi della schermata quando basta aggiornare il singolo elemento.",
        "Spostare caricamenti non essenziali dopo il primo render e misurare nuovamente il tempo di risposta."
      ],
      safeAutoRepair: false,
      repairNote: "Serve una correzione nel codice: nessuna modifica automatica viene applicata senza test."
    };

    if (impact === "commesse") return {
      title: "Protezione isolamento commessa",
      summary: "L'errore può coinvolgere associazioni fra commesse e impianti. La priorità è evitare qualsiasi scrittura o render con un commessaId non corrispondente.",
      steps: [
        "Verificare che ogni query/listener sia filtrato dal commessaId attivo.",
        "Scartare snapshot arrivati dopo il cambio commessa se appartengono alla commessa precedente.",
        "Prima del render, validare che ogni impianto appartenga realmente alla commessa aperta.",
        "Non eseguire correzioni automatiche sui dati finché la causa non è verificata."
      ],
      safeAutoRepair: false,
      repairNote: "Per sicurezza non viene mai modificato automaticamente alcun dato di commesse o impianti."
    };

    if (impact === "sync") return {
      title: "Ripristino salvataggio / sincronizzazione",
      summary: "Il problema riguarda Firestore, rete, offline o una scrittura non completata.",
      steps: [
        "Verificare permission-denied, rete e percorso Firestore esatto.",
        "Mantenere la modifica locale in coda senza annullare lo stato UI già confermato dall'utente.",
        "Ritentare la sincronizzazione in modo idempotente e impedire scritture duplicate."
      ],
      safeAutoRepair: false,
      repairNote: "Il Centro errori prepara la diagnosi ma non riscrive automaticamente dati operativi."
    };

    if (impact === "auth") return {
      title: "Sessione / permessi",
      summary: "Controllare sessione Firebase, ruolo utente e regole di accesso senza cancellare credenziali salvate.",
      steps: [
        "Verificare che Firebase sia inizializzato prima di usare auth/firestore.",
        "Controllare token e ruolo senza forzare logout automatici.",
        "Verificare regole e Cloud Functions interessate dal permission-denied."
      ],
      safeAutoRepair: false,
      repairNote: "Nessun logout o cambio permessi viene eseguito automaticamente."
    };

    if (impact === "map") return {
      title: "Mappa / GPS",
      summary: "Controllare coordinate, permessi di posizione e stato della mappa senza cancellare dati impianto.",
      steps: [
        "Validare latitudine e longitudine prima di creare marker o avviare la navigazione.",
        "Gestire il rifiuto del permesso GPS con fallback esplicito.",
        "Evitare ricreazioni complete della mappa che chiudono popup o bloccano i controlli."
      ],
      safeAutoRepair: false,
      repairNote: "Le coordinate salvate non vengono cambiate automaticamente."
    };

    if (impact === "blocking") return {
      title: "Sblocco app",
      summary: "Il problema impedisce o ostacola l'uso dell'app. Va isolata la funzione responsabile prima di qualunque modifica automatica.",
      steps: [
        "Controllare l'ultimo evento e l'azione immediatamente precedente al blocco.",
        "Verificare overlay, pointer-events, loop di render e task sincroni lunghi.",
        "Correggere su branch dedicato e testare Home, apertura commesse, FATTO e navigazione prima del merge."
      ],
      safeAutoRepair: false,
      repairNote: "Riparazione automatica disabilitata per evitare regressioni sui flussi critici."
    };

    return generic;
  }

  function githubRepairIssueUrl(item) {
    const impact = impactCategory(item);
    const diagnostic = {
      errorCenterId: item?.id || "",
      title: item?.title || "",
      category: item?.category || "",
      impact: impact?.label || "",
      severity: item?.severity || "",
      feature: item?.feature || "",
      page: item?.lastPage || item?.lastActiveView || "",
      platform: item?.lastPlatform || "",
      appVersion: item?.lastAppVersion || "",
      message: item?.lastMessage || "",
      stack: String(item?.lastStack || "").slice(0, 5000),
      commessaId: item?.commessaId || "",
      commessaName: item?.commessaName || "",
      impiantoId: item?.impiantoId || ""
    };
    const title = '[AUTO-REPAIR] ' + String(item?.title || item?.category || 'Errore app').slice(0, 160);
    const advice = repairAdvice(item);
    const body = [
      '## Richiesta automatica dal Centro errori',
      '',
      '**Obiettivo:** correggere questo errore con la modifica minima e sicura, senza alterare i flussi non coinvolti.',
      '',
      '**Impatto:** ' + (impact?.icon || '') + ' ' + (impact?.label || ''),
      '**Gravità:** ' + (item?.severity || ''),
      '',
      '### Diagnostica',
      '~~~json',
      JSON.stringify(diagnostic, null, 2),
      '~~~',
      '',
      '### Indicazione del Centro errori',
      advice.summary || '',
      '',
      '### Vincoli di sicurezza',
      '- Non modificare né indebolire la protezione FATTO / Whazzup.',
      '- Non modificare automaticamente dati di commesse o impianti.',
      '- Non eseguire merge automatico.',
      '- Creare una PR in bozza solo se i controlli critici passano.'
    ].join('\n');
    return 'https://github.com/ionut290/hera-app/issues/new?title=' + encodeURIComponent(title) + '&body=' + encodeURIComponent(body);
  }

  function startGithubRepair() {
    const item = selectedItem();
    if (!item) return;
    const root = document.querySelector('[data-error-repair-result]');
    const url = githubRepairIssueUrl(item);
    if (root) root.innerHTML = '<div class="hera-error-solution-note">🚀 Sto aprendo GitHub con la diagnostica già compilata. Dopo aver creato la richiesta, il workflow Codex prepara una correzione su branch separato, esegue i controlli critici e crea una PR in bozza solo se i test passano.</div>';
    const opened = window.open(url, '_blank', 'noopener,noreferrer');
    if (!opened) location.href = url;
  }

  function solutionHtml(item) {
    const advice = repairAdvice(item);
    const steps = advice.steps.map((step, index) => '<li><strong>' + (index + 1) + '.</strong> ' + esc(step) + '</li>').join('');
    return '<div class="hera-error-solution-card"><div class="hera-error-solution-title">🧠 ' + esc(advice.title) + '</div><p>' + esc(advice.summary) + '</p><ol>' + steps + '</ol><div class="hera-error-solution-note">' + esc(advice.repairNote) + '</div></div>';
  }

  function ensureImpactStyle() {
    if (document.getElementById("hera-error-impact-style")) return;
    const style = document.createElement("style");
    style.id = "hera-error-impact-style";
    style.textContent = ".hera-error-impact-summary{margin:10px 14px 0;padding:11px;border:1px solid #dbe3ef;border-radius:14px;background:#fff}.hera-error-impact-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:9px}.hera-error-impact-head strong{font-size:.82rem}.hera-error-impact-head small{color:#64748b;font-size:.68rem}.hera-error-impact-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:7px}.hera-error-impact-card{display:grid;grid-template-columns:auto 1fr auto;gap:7px;align-items:center;padding:9px;border:1px solid #e2e8f0;border-radius:11px;background:#f8fafc;color:#172033;text-align:left;cursor:pointer}.hera-error-impact-card:hover,.hera-error-impact-card.is-active{border-color:#93c5fd;background:#eff6ff}.hera-error-impact-card b{font-size:.75rem}.hera-error-impact-card span:last-child{min-width:28px;padding:3px 6px;border-radius:999px;background:#e2e8f0;font-size:.68rem;font-weight:900;text-align:center}.hera-error-impact-pill{display:inline-flex!important;width:max-content!important;margin-top:6px!important;padding:3px 7px;border-radius:999px;background:#eef2ff;color:#334155;font-size:.67rem!important;font-weight:800}.hera-error-tag[data-impact=blocking]{background:#fee2e2;color:#991b1b}.hera-error-tag[data-impact=commesse]{background:#fff1f2;color:#9f1239}.hera-error-tag[data-impact=performance]{background:#fef3c7;color:#92400e}.hera-error-tag[data-impact=sync]{background:#e0f2fe;color:#075985}.hera-error-tag[data-impact=notifications]{background:#f3e8ff;color:#6b21a8}.hera-error-impact-empty{padding:8px 2px;color:#64748b;font-size:.72rem}.hera-error-solution-card{margin:10px 0;padding:12px;border:1px solid #bfdbfe;border-radius:12px;background:#f8fbff}.hera-error-solution-title{font-weight:900;margin-bottom:6px}.hera-error-solution-card p{margin:0 0 8px}.hera-error-solution-card ol{margin:0;padding-left:21px}.hera-error-solution-card li{margin:6px 0}.hera-error-solution-note{margin-top:10px;padding:8px 10px;border-radius:9px;background:#eef2ff;font-size:.76rem}.hera-error-repair-actions{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:10px}.hera-error-repair-btn{font-weight:900}.hera-error-repair-btn.is-safe{background:#15803d;color:#fff;border-color:#15803d}.hera-error-repair-result{margin:8px 0 0}@media(max-width:880px){.hera-error-impact-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}@media(max-width:520px){.hera-error-impact-summary{margin:8px 8px 0}.hera-error-impact-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.hera-error-impact-card{padding:8px 7px}.hera-error-impact-card b{font-size:.7rem}.hera-error-impact-head{align-items:flex-start;flex-direction:column}}";
    document.head.appendChild(style);
  }

  function renderImpactSummary() {
    const root = document.querySelector("[data-error-impact-summary]");
    if (!root) return;
    ensureImpactStyle();
    if (!state.dataVerified && state.loading) {
      root.innerHTML = '<div class="hera-error-impact-empty">Classificazione degli errori in preparazione…</div>';
      return;
    }
    const counts = Object.fromEntries(IMPACT_CATEGORIES.map((category) => [category.id, 0]));
    state.items.forEach((item) => { counts[impactCategory(item).id] += Math.max(1, Number(item.occurrences || 0)); });
    const cards = IMPACT_CATEGORIES.map((category) => '<button type="button" class="hera-error-impact-card ' + (state.impactFilter === category.id ? 'is-active' : '') + '" data-error-impact="' + esc(category.id) + '" title="' + esc(category.note) + '"><span aria-hidden="true">' + category.icon + '</span><b>' + esc(category.label) + '</b><span>' + counts[category.id] + '</span></button>').join("");
    root.innerHTML = '<div class="hera-error-impact-head"><div><strong>Tipo di impatto</strong><br><small>Classificazione automatica delle occorrenze caricate</small></div><button type="button" class="btn ' + (state.impactFilter === 'all' ? 'btn-primary' : '') + '" data-error-impact="all">TUTTI</button></div><div class="hera-error-impact-grid">' + cards + '</div>';
  }

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
      ["Aperti", counts.open, ""],
      ["In verifica", counts.inVerification, ""],
      ["Critici", counts.critical, "is-critical"],
      ["Alti", counts.high, "is-high"],
      ["Risolti", counts.resolved, ""],
      ["Ignorati", counts.ignored, ""]
    ].map(([label, value, className]) => {
      const display = state.dataVerified && Number.isFinite(Number(value)) ? Number(value) : "—";
      return `<div class="hera-error-summary-card ${className}"><strong>${display}</strong><span>${esc(label)}</span></div>`;
    }).join("");
  }

  function renderHealth() {
    const root = document.querySelector("[data-error-health]");
    if (!root) return;
    const monitor = window.HeraAppErrorMonitor;
    const health = monitor?.getHealth?.() || {};
    const queue = Math.max(0, Number(health.queuedReports || monitor?.getQueueLength?.() || 0));
    if (state.loading) {
      root.className = "hera-error-health is-checking";
      root.textContent = "⏳ Verifica collegamento con il server…";
      return;
    }
    if (state.loadError) {
      root.className = "hera-error-health is-error";
      root.innerHTML = `<strong>❌ Dati non verificati.</strong> ${esc(state.loadError)}${queue ? ` · ${queue} segnalazioni in attesa sul dispositivo.` : ""}`;
      return;
    }
    if (state.dataVerified) {
      root.className = "hera-error-health is-ok";
      root.innerHTML = `<strong>✅ Collegato a Firestore.</strong> Contatori verificati dal server ${esc(formatDate(state.serverTime))}.${queue ? ` · ${queue} segnalazioni in attesa di invio.` : " · Nessuna segnalazione in attesa."}${state.truncated ? " · Elenco limitato agli errori più recenti; i contatori restano completi." : ""}`;
      return;
    }
    root.className = "hera-error-health is-checking";
    root.textContent = "Contatori non ancora verificati.";
  }

  function renderList() {
    const root = document.querySelector("[data-error-list]");
    if (!root) return;
    if (state.loading) {
      root.innerHTML = '<div class="hera-error-loading">Caricamento errori…</div>';
      return;
    }
    if (state.loadError) {
      root.innerHTML = `<div class="hera-error-empty"><strong>Centro errori non disponibile.</strong><br>${esc(state.loadError)}<br><small>Gli zeri non vengono mostrati perché non sono stati verificati dal server.</small></div>`;
      return;
    }
    const visibleItems = impactItems();
    if (!visibleItems.length) {
      root.innerHTML = '<div class="hera-error-empty"><strong>Nessun errore trovato.</strong><br>Modifica i filtri oppure la categoria di impatto.</div>';
      return;
    }
    root.innerHTML = visibleItems.map((item) => `
      <button type="button" class="hera-error-row ${state.selectedId === item.id ? "is-selected" : ""}" data-error-id="${esc(item.id)}">
        <span class="hera-error-dot" data-severity="${esc(item.severity)}"></span>
        <span class="hera-error-row-copy"><strong>${esc(item.title || item.category || "Errore app")}</strong><span>${esc(item.lastMessage || "Nessun messaggio")}</span><small>${esc(statusLabel(item.status))} · ${esc(formatDate(item.lastSeenAt))} · ${esc(item.lastPlatform || "dispositivo non noto")}</small><span class="hera-error-impact-pill">${impactCategory(item).icon} ${esc(impactCategory(item).label)}</span></span>
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
      <div class="hera-error-tags"><span class="hera-error-tag" data-severity="${esc(item.severity)}">${esc(severityLabel(item.severity))}</span><span class="hera-error-tag" data-impact="${esc(impactCategory(item).id)}">${impactCategory(item).icon} ${esc(impactCategory(item).label)}</span><span class="hera-error-tag">${esc(statusLabel(item.status))}</span><span class="hera-error-tag">${Number(item.occurrences || 0)} occorrenze</span><span class="hera-error-tag">${Number(item.affectedUsers || 0)} utenti</span></div>
      <dl class="hera-error-detail-grid">
        <div><dt>Funzione</dt><dd>${esc(item.feature || "—")}</dd></div><div><dt>Ultima schermata</dt><dd>${esc(item.lastActiveView || item.lastPage || "—")}</dd></div>
        <div><dt>Prima comparsa</dt><dd>${esc(formatDate(item.firstSeenAt))}</dd></div><div><dt>Ultima comparsa</dt><dd>${esc(formatDate(item.lastSeenAt))}</dd></div>
        <div><dt>Dispositivo</dt><dd>${esc(item.lastPlatform || "—")}</dd></div><div><dt>Versione app</dt><dd>${esc(item.lastAppVersion || "—")}</dd></div>
        <div><dt>Utente</dt><dd>${esc(item.lastUserName || item.lastUserEmailMasked || "—")}</dd></div><div><dt>Durata / tocchi</dt><dd>${item.lastDurationMs ? `${Number(item.lastDurationMs)} ms` : "—"}${item.lastTapCount ? ` · ${Number(item.lastTapCount)} tocchi` : ""}</dd></div>
        <div><dt>Commessa</dt><dd>${esc(item.commessaName || item.commessaId || "—")}</dd></div><div><dt>Impianto</dt><dd>${esc(item.impiantoId || "—")}</dd></div>
      </dl>
      <h4>Impatto operativo</h4><div class="hera-error-action"><strong>${impactCategory(item).icon} ${esc(impactCategory(item).label)}</strong><br>${esc(impactCategory(item).note)}</div>
      <h4>Problema</h4><div class="hera-error-message">${esc(item.lastMessage || "Nessun messaggio")}</div>
      <h4>Azione tecnica consigliata</h4><div class="hera-error-action">${esc(item.diagnosisAction || "Analizzare il dettaglio tecnico.")}</div>
      <h4>Stack tecnico</h4><pre class="hera-error-stack">${esc(item.lastStack || "Non disponibile")}</pre>
      <h4>Ultimi eventi registrati</h4><div class="hera-error-events">${(item.recentEvents || []).map(eventDetails).join("") || '<p class="muted">Nessun evento dettagliato disponibile.</p>'}</div>
      <section class="hera-error-admin">
        <label>Stato<select data-error-detail-status>${["open", "in_verification", "resolved", "ignored"].map((status) => `<option value="${status}" ${item.status === status ? "selected" : ""}>${esc(statusLabel(status))}</option>`).join("")}</select></label>
        <label>Nota amministratore<textarea class="hera-error-note" data-error-note rows="4" maxlength="3000" placeholder="Annota verifica, causa o correzione applicata.">${esc(item.adminNote || "")}</textarea></label>
        <div class="hera-error-repair-actions"><button class="btn hera-error-repair-btn" type="button" data-error-find-solution>🧠 TROVA SOLUZIONE</button><button class="btn hera-error-repair-btn ${repairAdvice(item).safeAutoRepair ? "is-safe" : ""}" type="button" data-error-repair>🛠️ RIPARA ERRORE</button><button class="btn hera-error-repair-btn" type="button" data-error-github-repair>🚀 RIPARA SU GITHUB</button><button class="btn hera-error-repair-btn hera-error-chatgpt-btn" type="button" data-error-chatgpt-category>🤖 INVIA CATEGORIA A CHATGPT</button></div>
        <div class="hera-error-repair-result" data-error-repair-result></div>
        <p class="hera-error-status" data-error-feedback></p>
        <div class="hera-error-admin-actions"><button class="btn btn-primary" type="button" data-error-save>SALVA STATO</button><button class="btn" type="button" data-error-copy>COPIA DIAGNOSTICA</button></div>
      </section>`;
  }

  async function loadDashboard() {
    if (!state.admin || state.loading) return false;
    let loaded = false;
    state.loading = true;
    state.dataVerified = false;
    state.loadError = "";
    renderHealth();
    renderSummary();
    renderImpactSummary();
    renderList();
    const query = document.querySelector("[data-error-query]")?.value || "";
    const status = document.querySelector("[data-error-status]")?.value || "all";
    const severity = document.querySelector("[data-error-severity]")?.value || "all";
    try {
      const result = await callFunction(FUNCTIONS.dashboard, { query, status, severity, limit: 120 });
      if (!result || result.countsVerified !== true || !result.counts || !Array.isArray(result.items)) {
        throw new Error("Il server non ha certificato i contatori. È necessario completare il deploy del Centro errori.");
      }
      state.items = Array.isArray(result.items) ? result.items : [];
      state.counts = result.counts || {};
      state.dataVerified = true;
      state.serverTime = result.serverTime || new Date().toISOString();
      state.truncated = Boolean(result.truncated);
      loaded = true;
      if (state.pendingTarget && state.items.some((item) => item.id === state.pendingTarget)) state.selectedId = state.pendingTarget;
      else if (!state.items.some((item) => item.id === state.selectedId)) state.selectedId = state.items[0]?.id || "";
      state.pendingTarget = "";
    } catch (error) {
      console.error("Caricamento Centro errori fallito:", error);
      state.items = [];
      state.counts = {};
      state.dataVerified = false;
      state.serverTime = "";
      state.truncated = false;
      state.loadError = error?.message || "Controlla il deploy delle Cloud Functions.";
    } finally {
      state.loading = false;
      renderHealth();
      renderSummary();
      renderImpactSummary();
      renderList();
      renderDetail();
    }
    return loaded;
  }

  async function markSeen() {
    try {
      await callFunction(FUNCTIONS.seen);
      setBadge(0);
    } catch (error) {
      console.warn("Conferma lettura Centro errori non sincronizzata:", error);
    }
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
    const loaded = await loadDashboard();
    if (loaded) void markSeen();
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

  async function sendCategoryToChatGpt() {
    const feedback = document.querySelector("[data-error-feedback]");
    const items = chatGptCategoryItems();
    if (!items.length) {
      if (feedback) feedback.textContent = "⚠️ Nessun errore disponibile nella categoria selezionata.";
      return;
    }
    const chatWindow = window.open("https://chatgpt.com/", "_blank");
    try { if (chatWindow) chatWindow.opener = null; } catch (_) {}
    const text = chatGptCategoryText();
    try {
      if (navigator.clipboard?.writeText) await navigator.clipboard.writeText(text);
      else {
        const textarea = document.createElement("textarea");
        textarea.value = text;
        textarea.style.position = "fixed";
        textarea.style.opacity = "0";
        document.body.appendChild(textarea);
        textarea.select();
        if (!document.execCommand("copy")) throw new Error("Copia non supportata");
        textarea.remove();
      }
      if (feedback) feedback.textContent = `✅ ${items.length} gruppi della categoria copiati. In ChatGPT premi Incolla e Invia.`;
      if (!chatWindow) window.alert("Blocco completo copiato. Apri ChatGPT e premi Incolla.");
    } catch (error) {
      if (chatWindow) chatWindow.close?.();
      if (feedback) feedback.textContent = `⚠️ Copia non riuscita: ${error?.message || "errore sconosciuto"}. Usa COPIA DIAGNOSTICA per il singolo errore.`;
    }
  }

  function showSelectedSolution() {
    const item = selectedItem();
    const root = document.querySelector("[data-error-repair-result]");
    if (!item || !root) return;
    root.innerHTML = solutionHtml(item);
  }

  async function repairSelectedError() {
    const item = selectedItem();
    const root = document.querySelector("[data-error-repair-result]");
    const feedback = document.querySelector("[data-error-feedback]");
    if (!item || !root) return;
    const advice = repairAdvice(item);
    root.innerHTML = solutionHtml(item);

    if (!advice.safeAutoRepair) {
      if (feedback) feedback.textContent = "🛡️ Riparazione automatica bloccata: serve una modifica tecnica verificata. La soluzione guidata è pronta sopra.";
      return;
    }

    const confirmed = window.confirm("Questa riparazione agisce solo sulla configurazione locale sicura del dispositivo e non modifica commesse, impianti, FATTO o WHAZZUP. Procedere?");
    if (!confirmed) return;

    if (/applicationserverkey|p-256|vapid/.test(normalizeImpactText(item))) {
      try {
        localStorage.setItem("hera:webPushAutoDisabled", "1");
        localStorage.setItem("hera:disableWebPush", "1");
      } catch (_) {}
      try {
        const registration = await navigator.serviceWorker?.ready;
        const subscription = await registration?.pushManager?.getSubscription?.();
        if (subscription) await subscription.unsubscribe();
      } catch (error) {
        console.warn("Pulizia subscription Web Push non completata:", error);
      }
      if (feedback) feedback.textContent = "✅ Tentativi Web Push automatici disattivati su questo dispositivo e subscription locale rimossa. Ora verifica se l'errore ricompare.";
      return;
    }

    if (feedback) feedback.textContent = "ℹ️ Nessuna azione automatica disponibile per questo errore.";
  }

  function handleCenterClick(event) {
    if (event.target.closest?.("[data-error-close]")) return closeCenter();
    if (event.target.closest?.("[data-error-refresh], [data-error-apply]")) return void loadDashboard();
    const impactButton = event.target.closest?.("[data-error-impact]");
    if (impactButton) {
      state.impactFilter = impactButton.dataset.errorImpact || "all";
      const visible = impactItems();
      if (!visible.some((item) => item.id === state.selectedId)) state.selectedId = visible[0]?.id || "";
      renderImpactSummary();
      renderList();
      renderDetail();
      return;
    }
    const row = event.target.closest?.("[data-error-id]");
    if (row) {
      state.selectedId = row.dataset.errorId || "";
      renderList();
      renderDetail();
      return;
    }
    if (event.target.closest?.("[data-error-find-solution]")) return showSelectedSolution();
    if (event.target.closest?.("[data-error-repair]")) return void repairSelectedError();
    if (event.target.closest?.("[data-error-github-repair]")) return startGithubRepair();
    if (event.target.closest?.("[data-error-chatgpt-category]")) return void sendCategoryToChatGpt();
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

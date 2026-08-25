const fs = require('fs');
const path = require('path');

const file = path.join(__dirname, '..', 'admin-error-center.js');
let source = fs.readFileSync(file, 'utf8');

function replaceOnce(label, from, to) {
  if (!source.includes(from)) throw new Error(`Patch ${label}: anchor not found`);
  source = source.replace(from, to);
}

replaceOnce('version', 'const VERSION = "1.1.0";', 'const VERSION = "1.2.0";');

replaceOnce(
  'state impact filter',
  '    serverTime: "",\n    truncated: false\n',
  '    serverTime: "",\n    truncated: false,\n    impactFilter: "all"\n'
);

replaceOnce(
  'impact section markup',
  '          <section class="hera-error-summary" data-error-summary></section>\n          <section class="hera-error-toolbar">',
  '          <section class="hera-error-summary" data-error-summary></section>\n          <section class="hera-error-impact-summary" data-error-impact-summary aria-label="Categorie di impatto degli errori"></section>\n          <section class="hera-error-toolbar">'
);

replaceOnce(
  'impact classifier',
  '  const statusLabel = (status) => ({ open: "Aperto", in_verification: "In verifica", resolved: "Risolto", ignored: "Ignorato" })[status] || "Aperto";\n\n  function formatDate(value) {',
  String.raw`  const statusLabel = (status) => ({ open: "Aperto", in_verification: "In verifica", resolved: "Risolto", ignored: "Ignorato" })[status] || "Aperto";

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

    if (match(/\b(crash|fatal|freeze|frozen|hang|hung|unresponsive|uncaught)\b|blocca(?:to|ta|re)?|bloccato|schermata bianca|non risponde|non reagisce|maximum call stack|out of memory/)) {
      return IMPACT_CATEGORIES[0];
    }
    if (match(/commessa|impianto/) && match(/mescol|sovrascr|perdit|pers[oi]|cancell|elimin|spostat|assegnat.*sbagli|id errat|integrit|isolament|altra commessa|commessa sbagliata|impianti di un altra|impianti di un'altra/)) {
      return IMPACT_CATEGORIES[1];
    }
    if (match(/notific|notification|push|messaging|fcm|vapid|applicationserverkey|p-256|service.?worker.*push/)) {
      return IMPACT_CATEGORIES[5];
    }
    if (match(/mappa|\bmap\b|gps|geoloc|leaflet|google maps|coordinate|latitud|longitud|navigaz/)) {
      return IMPACT_CATEGORIES[6];
    }
    if (match(/\bauth\b|login|permission-denied|permess|unauthor|forbidden|credential|sessione|token.*scad|accesso negato/)) {
      return IMPACT_CATEGORIES[4];
    }
    if (match(/rallent|\bslow\b|performance|long task|listener duplic|listener.*multip|render.*ripet|memory leak|troppe letture|richieste ripet|latency|latenza|timeout/)) {
      return IMPACT_CATEGORIES[2];
    }
    if (match(/firestore|database|storage|sincron|\bsync\b|offline|salvat|scritt|\bwrite\b|network|rete|quota/)) {
      return IMPACT_CATEGORIES[3];
    }
    return IMPACT_CATEGORIES[7];
  }

  function impactItems() {
    return state.impactFilter === "all"
      ? state.items
      : state.items.filter((item) => impactCategory(item).id === state.impactFilter);
  }

  function ensureImpactStyle() {
    if (document.getElementById("hera-error-impact-style")) return;
    const style = document.createElement("style");
    style.id = "hera-error-impact-style";
    style.textContent = ".hera-error-impact-summary{margin:10px 14px 0;padding:11px;border:1px solid #dbe3ef;border-radius:14px;background:#fff}.hera-error-impact-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:9px}.hera-error-impact-head strong{font-size:.82rem}.hera-error-impact-head small{color:#64748b;font-size:.68rem}.hera-error-impact-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:7px}.hera-error-impact-card{display:grid;grid-template-columns:auto 1fr auto;gap:7px;align-items:center;padding:9px;border:1px solid #e2e8f0;border-radius:11px;background:#f8fafc;color:#172033;text-align:left;cursor:pointer}.hera-error-impact-card:hover,.hera-error-impact-card.is-active{border-color:#93c5fd;background:#eff6ff}.hera-error-impact-card b{font-size:.75rem}.hera-error-impact-card span:last-child{min-width:28px;padding:3px 6px;border-radius:999px;background:#e2e8f0;font-size:.68rem;font-weight:900;text-align:center}.hera-error-impact-pill{display:inline-flex!important;width:max-content!important;margin-top:6px!important;padding:3px 7px;border-radius:999px;background:#eef2ff;color:#334155;font-size:.67rem!important;font-weight:800}.hera-error-tag[data-impact=blocking]{background:#fee2e2;color:#991b1b}.hera-error-tag[data-impact=commesse]{background:#fff1f2;color:#9f1239}.hera-error-tag[data-impact=performance]{background:#fef3c7;color:#92400e}.hera-error-tag[data-impact=sync]{background:#e0f2fe;color:#075985}.hera-error-tag[data-impact=notifications]{background:#f3e8ff;color:#6b21a8}.hera-error-impact-empty{padding:8px 2px;color:#64748b;font-size:.72rem}@media(max-width:880px){.hera-error-impact-grid{grid-template-columns:repeat(2,minmax(0,1fr))}}@media(max-width:520px){.hera-error-impact-summary{margin:8px 8px 0}.hera-error-impact-grid{grid-template-columns:repeat(2,minmax(0,1fr))}.hera-error-impact-card{padding:8px 7px}.hera-error-impact-card b{font-size:.7rem}.hera-error-impact-head{align-items:flex-start;flex-direction:column}}";
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
    state.items.forEach((item) => {
      const category = impactCategory(item);
      counts[category.id] += Math.max(1, Number(item.occurrences || 0));
    });
    const cards = IMPACT_CATEGORIES.map((category) => `
      <button type="button" class="hera-error-impact-card ${state.impactFilter === category.id ? "is-active" : ""}" data-error-impact="${esc(category.id)}" title="${esc(category.note)}">
        <span aria-hidden="true">${category.icon}</span><b>${esc(category.label)}</b><span>${counts[category.id]}</span>
      </button>`).join("");
    root.innerHTML = `<div class="hera-error-impact-head"><div><strong>Tipo di impatto</strong><br><small>Classificazione automatica delle occorrenze caricate</small></div><button type="button" class="btn ${state.impactFilter === "all" ? "btn-primary" : ""}" data-error-impact="all">TUTTI</button></div><div class="hera-error-impact-grid">${cards}</div>`;
  }

  function formatDate(value) {`
);

replaceOnce(
  'list visibility',
  '    if (!state.items.length) {\n      root.innerHTML = \'<div class="hera-error-empty"><strong>Nessun errore trovato.</strong><br>Modifica i filtri oppure aggiorna il pannello.</div>\';\n      return;\n    }\n    root.innerHTML = state.items.map((item) => `',
  '    const visibleItems = impactItems();\n    if (!visibleItems.length) {\n      root.innerHTML = \'<div class="hera-error-empty"><strong>Nessun errore trovato.</strong><br>Modifica i filtri oppure la categoria di impatto.</div>\';\n      return;\n    }\n    root.innerHTML = visibleItems.map((item) => `'
);

replaceOnce(
  'row impact pill',
  '<span class="hera-error-row-copy"><strong>${esc(item.title || item.category || "Errore app")}</strong><span>${esc(item.lastMessage || "Nessun messaggio")}</span><small>${esc(statusLabel(item.status))} · ${esc(formatDate(item.lastSeenAt))} · ${esc(item.lastPlatform || "dispositivo non noto")}</small></span>',
  '<span class="hera-error-row-copy"><strong>${esc(item.title || item.category || "Errore app")}</strong><span>${esc(item.lastMessage || "Nessun messaggio")}</span><small>${esc(statusLabel(item.status))} · ${esc(formatDate(item.lastSeenAt))} · ${esc(item.lastPlatform || "dispositivo non noto")}</small><span class="hera-error-impact-pill">${impactCategory(item).icon} ${esc(impactCategory(item).label)}</span></span>'
);

replaceOnce(
  'detail impact tag',
  '<div class="hera-error-tags"><span class="hera-error-tag" data-severity="${esc(item.severity)}">${esc(severityLabel(item.severity))}</span><span class="hera-error-tag">${esc(statusLabel(item.status))}</span><span class="hera-error-tag">${Number(item.occurrences || 0)} occorrenze</span><span class="hera-error-tag">${Number(item.affectedUsers || 0)} utenti</span></div>',
  '<div class="hera-error-tags"><span class="hera-error-tag" data-severity="${esc(item.severity)}">${esc(severityLabel(item.severity))}</span><span class="hera-error-tag" data-impact="${esc(impactCategory(item).id)}">${impactCategory(item).icon} ${esc(impactCategory(item).label)}</span><span class="hera-error-tag">${esc(statusLabel(item.status))}</span><span class="hera-error-tag">${Number(item.occurrences || 0)} occorrenze</span><span class="hera-error-tag">${Number(item.affectedUsers || 0)} utenti</span></div>'
);

replaceOnce(
  'detail impact explanation',
  '      <h4>Problema</h4><div class="hera-error-message">${esc(item.lastMessage || "Nessun messaggio")}</div>\n      <h4>Azione tecnica consigliata</h4>',
  '      <h4>Impatto operativo</h4><div class="hera-error-action"><strong>${impactCategory(item).icon} ${esc(impactCategory(item).label)}</strong><br>${esc(impactCategory(item).note)}</div>\n      <h4>Problema</h4><div class="hera-error-message">${esc(item.lastMessage || "Nessun messaggio")}</div>\n      <h4>Azione tecnica consigliata</h4>'
);

replaceOnce(
  'loading impact render',
  '    renderHealth();\n    renderSummary();\n    renderList();',
  '    renderHealth();\n    renderSummary();\n    renderImpactSummary();\n    renderList();'
);

replaceOnce(
  'final impact render',
  '      renderHealth();\n      renderSummary();\n      renderList();\n      renderDetail();',
  '      renderHealth();\n      renderSummary();\n      renderImpactSummary();\n      renderList();\n      renderDetail();'
);

replaceOnce(
  'impact click handler',
  '  function handleCenterClick(event) {\n    if (event.target.closest?.("[data-error-close]")) return closeCenter();\n    if (event.target.closest?.("[data-error-refresh], [data-error-apply]")) return void loadDashboard();\n    const row = event.target.closest?.("[data-error-id]");',
  '  function handleCenterClick(event) {\n    if (event.target.closest?.("[data-error-close]")) return closeCenter();\n    if (event.target.closest?.("[data-error-refresh], [data-error-apply]")) return void loadDashboard();\n    const impactButton = event.target.closest?.("[data-error-impact]");\n    if (impactButton) {\n      state.impactFilter = impactButton.dataset.errorImpact || "all";\n      const visible = impactItems();\n      if (!visible.some((item) => item.id === state.selectedId)) state.selectedId = visible[0]?.id || "";\n      renderImpactSummary();\n      renderList();\n      renderDetail();\n      return;\n    }\n    const row = event.target.closest?.("[data-error-id]");'
);

fs.writeFileSync(file, source, 'utf8');
console.log('Applied Centro errori impact categories patch.');

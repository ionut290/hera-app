(() => {
  'use strict';

  if (window.HeraControlCenterOrganizer?.installed) return;

  const RECENTS_KEY = 'heraControlCenterRecentV2';
  const MAX_RECENTS = 5;
  const GROUPS = [
    { id: 'operations', icon: '🧭', title: 'Operatività', description: 'Commesse, impianti, squadre e attività giornaliere.', keywords: ['commessa','commesse','impianto','impianti','squadra','squadre','ore','programmazione','operativ','lavoro','attivita'] },
    { id: 'people', icon: '👥', title: 'Persone e accessi', description: 'Utenti, personale, mezzi, account e permessi.', keywords: ['utente','utenti','personale','accesso','accessi','account','password','profilo','mezzi','mezzo','operatore'] },
    { id: 'communication', icon: '🔔', title: 'Comunicazioni', description: 'Notifiche, banner, avvisi e segnalazioni.', keywords: ['notific','banner','avvisi','segnalaz','comunicaz','informazioni utili','messaggi'] },
    { id: 'documents', icon: '📁', title: 'Documenti e dati', description: 'Documenti, esportazioni, backup e archivi.', keywords: ['document','pdf','pos','drive','backup','archivio','esporta','importa','dati'] },
    { id: 'performance', icon: '⚡', title: 'Firestore e prestazioni', description: 'Consumi, diagnostica, cache e velocità.', keywords: ['firestore','letture','scritture','listener','prestaz','performance','diagnostic','cache','storage','quota','veloc','consumo','memoria'] },
    { id: 'security', icon: '🛡️', title: 'Sicurezza e manutenzione', description: 'Controlli, protezioni, ripristino e pulizia.', keywords: ['sicurezza','protezione','ripristino','manutenz','pulizia','reset','errore','log','verifica'] },
    { id: 'other', icon: '🧰', title: 'Altri strumenti', description: 'Strumenti amministrativi non classificati.', keywords: [] }
  ];

  const state = { root: null, host: null, observer: null, cards: [], activeGroup: 'all' };

  function normalize(value) {
    return String(value || '').toLocaleLowerCase('it-IT').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  function findRoot() {
    return document.getElementById('control-center-content') || document.getElementById('control-center-page');
  }

  function cardTitle(card) {
    return String(card?.querySelector('summary')?.textContent || card?.getAttribute('aria-label') || '').replace(/\s+/g, ' ').trim();
  }

  function cardSearchText(card) {
    return normalize(`${cardTitle(card)} ${card?.textContent || ''} ${card?.dataset.ccGroupTitle || ''}`);
  }

  function classify(card) {
    const haystack = cardSearchText(card);
    let best = GROUPS[GROUPS.length - 1];
    let score = 0;
    GROUPS.slice(0, -1).forEach((group) => {
      const current = group.keywords.reduce((total, keyword) => total + (haystack.includes(normalize(keyword)) ? 1 : 0), 0);
      if (current > score) { score = current; best = group; }
    });
    return best;
  }

  function safeReadRecents() {
    try {
      const value = JSON.parse(localStorage.getItem(RECENTS_KEY) || '[]');
      return Array.isArray(value) ? value.filter((item) => item && item.title).slice(0, MAX_RECENTS) : [];
    } catch (_) { return []; }
  }

  function saveRecent(card) {
    const title = cardTitle(card);
    if (!title) return;
    try {
      const groupId = card.dataset.ccGroup || 'other';
      const current = safeReadRecents().filter((item) => normalize(item.title) !== normalize(title));
      current.unshift({ title, groupId, usedAt: Date.now() });
      localStorage.setItem(RECENTS_KEY, JSON.stringify(current.slice(0, MAX_RECENTS)));
      renderRecents();
    } catch (_) {}
  }

  function ensureStyles() {
    if (document.getElementById('control-center-organizer-style')) return;
    const style = document.createElement('style');
    style.id = 'control-center-organizer-style';
    style.textContent = `
      .control-center-organized{--cc-line:#d7e3ed;--cc-ink:#17263b;--cc-muted:#607286;--cc-blue:#2563eb}
      .cc-organizer{display:grid;gap:10px;margin:6px 0 18px}
      .cc-dashboard{display:grid;gap:10px;padding:12px;border:1px solid #cfdeea;border-radius:18px;background:linear-gradient(145deg,#eaf3ff 0%,#edf8f3 58%,#f4effc 100%);box-shadow:0 8px 22px rgba(30,64,100,.07)}
      .cc-dashboard-top{display:flex;align-items:flex-start;justify-content:space-between;gap:10px}.cc-dashboard-title{margin:0;color:var(--cc-ink);font-size:1rem}.cc-dashboard-subtitle{margin:2px 0 0;color:var(--cc-muted);font-size:.78rem;line-height:1.35}
      .cc-dashboard-count{display:inline-flex;align-items:center;justify-content:center;min-width:42px;height:30px;padding:0 9px;border-radius:999px;background:#dceaff;color:#24528d;font-weight:900;font-size:.76rem;white-space:nowrap}
      .cc-organizer-tools{position:sticky;top:3px;z-index:12;display:grid;grid-template-columns:minmax(0,1fr) auto;gap:7px;padding:7px;background:rgba(239,246,252,.96);backdrop-filter:blur(14px);border:1px solid #d1dfeb;border-radius:15px;box-shadow:0 6px 18px rgba(30,64,100,.08)}
      .cc-search-wrap{position:relative}.cc-search-icon{position:absolute;left:12px;top:50%;transform:translateY(-50%);pointer-events:none}.cc-organizer-search{min-height:44px;width:100%;padding:9px 12px 9px 38px;border:1px solid #bfcfdd;border-radius:12px;background:#fff;font-size:16px;color:#17263b}.cc-organizer-search:focus{outline:3px solid rgba(37,99,235,.12);border-color:#78a5ee}
      .cc-organizer-clear{min-width:44px;min-height:44px;border:1px solid #cbd8e5;border-radius:12px;background:#f8fbff;font-weight:900;color:#41556d}.cc-organizer-hint{grid-column:1/-1;margin:0 3px;color:#64748b;font-size:.7rem}
      .cc-quick-actions{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:6px}.cc-quick-btn{min-height:38px;border:1px solid #cedbe7;border-radius:11px;background:rgba(255,255,255,.8);font-size:.72rem;font-weight:850;color:#31465e}.cc-quick-btn:active{transform:scale(.98)}
      .cc-filter-strip{display:flex;gap:6px;overflow-x:auto;padding:1px 0 3px;scrollbar-width:none}.cc-filter-strip::-webkit-scrollbar{display:none}.cc-filter-chip{flex:0 0 auto;min-height:34px;padding:6px 10px;border:1px solid #cfdae6;border-radius:999px;background:#f8fbfd;color:#42566d;font-size:.72rem;font-weight:800}.cc-filter-chip.is-active{background:#dceaff;border-color:#a9c7ef;color:#194f96}
      .cc-recents{display:none}.cc-recents.has-items{display:block}.cc-recents-head{display:flex;align-items:center;justify-content:space-between;gap:8px;margin:1px 2px 6px}.cc-recents-title{margin:0;color:#3b4f66;font-size:.74rem;text-transform:uppercase;letter-spacing:.04em}.cc-recents-list{display:flex;gap:6px;overflow-x:auto;scrollbar-width:none}.cc-recents-list::-webkit-scrollbar{display:none}.cc-recent-btn{flex:0 0 auto;max-width:210px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;padding:7px 10px;border:1px solid #d1deea;border-radius:10px;background:#fff;color:#334a62;font-size:.72rem;font-weight:750}.cc-recents-clear{border:0;background:transparent;color:#64748b;font-size:.7rem;text-decoration:underline}
      .cc-group{border:1px solid var(--cc-line);border-radius:16px;background:rgba(246,249,252,.94);overflow:hidden;box-shadow:0 4px 13px rgba(15,23,42,.04)}
      .cc-group-head{width:100%;display:grid;grid-template-columns:auto minmax(0,1fr) auto auto;align-items:center;gap:9px;min-height:50px;padding:9px 12px;border:0;background:linear-gradient(90deg,#eef4fa,#f4f8fa);color:var(--cc-ink);text-align:left}.cc-group-icon{font-size:1.15rem}.cc-group-copy{min-width:0}.cc-group-title{display:block;font-weight:900;font-size:.88rem}.cc-group-description{display:block;margin-top:1px;color:#718196;font-size:.66rem;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.cc-group-count{display:inline-grid;place-items:center;min-width:27px;height:27px;padding:0 7px;border-radius:999px;background:#dce8f4;color:#315477;font-size:.72rem;font-weight:900}.cc-group-chevron{font-size:1rem;transition:transform .16s ease}.cc-group.is-collapsed .cc-group-chevron{transform:rotate(-90deg)}
      .cc-group-body{display:grid;gap:4px;padding:4px}.cc-group.is-collapsed .cc-group-body{display:none}.cc-group-body>details{margin:0!important;border:1px solid #dde6ee!important;border-radius:12px!important;box-shadow:none!important;background:#fff!important;overflow:hidden}.cc-group-body>details>summary{min-height:43px;padding:8px 10px!important;font-weight:800;font-size:.82rem;background:#fff}.cc-group-body>details[open]>summary{background:#f5f9fc}.cc-group-body>details>*:not(summary){margin-left:9px!important;margin-right:9px!important}
      .cc-card-search-hidden,.cc-group-search-hidden,.cc-group-filter-hidden{display:none!important}.cc-no-results{display:none;padding:20px 14px;border:1px dashed #c8d4df;border-radius:14px;background:#f8fafc;text-align:center;color:#64748b}.cc-no-results.is-visible{display:block}.cc-no-results strong{display:block;color:#34475d;margin-bottom:3px}
      @media(min-width:900px){.cc-organizer{grid-template-columns:repeat(2,minmax(0,1fr));align-items:start}.cc-dashboard,.cc-organizer-tools,.cc-recents,.cc-filter-strip,.cc-no-results{grid-column:1/-1}.cc-group{height:max-content}}
      @media(max-width:700px){.cc-organizer{gap:7px;margin-top:3px}.cc-dashboard{padding:9px;border-radius:15px}.cc-dashboard-subtitle{font-size:.72rem}.cc-organizer-tools{top:2px;padding:5px;border-radius:13px}.cc-organizer-search,.cc-organizer-clear{min-height:41px;height:41px}.cc-quick-actions{gap:4px}.cc-quick-btn{min-height:35px;font-size:.66rem;padding:5px 3px}.cc-filter-chip{min-height:31px;padding:5px 8px;font-size:.67rem}.cc-group{border-radius:13px}.cc-group-head{min-height:45px;padding:7px 9px;gap:7px}.cc-group-title{font-size:.82rem}.cc-group-description{font-size:.61rem}.cc-group-body{gap:2px;padding:3px}.cc-group-body>details>summary{min-height:40px;padding:7px 8px!important;font-size:.77rem}}
    `;
    document.head.appendChild(style);
  }

  function collectCards(root) {
    return Array.from(root.querySelectorAll('details')).filter((card) => {
      if (card.closest('.cc-group')) return false;
      return !card.parentElement?.closest('details');
    });
  }

  function buildGroup(group) {
    const section = document.createElement('section');
    const initiallyOpen = group.id === 'performance' || group.id === 'operations';
    section.className = `cc-group${initiallyOpen ? '' : ' is-collapsed'}`;
    section.dataset.ccGroup = group.id;
    section.innerHTML = `<button class="cc-group-head" type="button" aria-expanded="${String(initiallyOpen)}"><span class="cc-group-icon" aria-hidden="true">${group.icon}</span><span class="cc-group-copy"><span class="cc-group-title">${group.title}</span><span class="cc-group-description">${group.description}</span></span><span class="cc-group-count">0</span><span class="cc-group-chevron" aria-hidden="true">⌄</span></button><div class="cc-group-body"></div>`;
    const button = section.querySelector('.cc-group-head');
    button.addEventListener('click', () => {
      const collapsed = section.classList.toggle('is-collapsed');
      button.setAttribute('aria-expanded', String(!collapsed));
    });
    return section;
  }

  function findCardByTitle(title) {
    const wanted = normalize(title);
    return state.cards.find((card) => normalize(cardTitle(card)) === wanted) || null;
  }

  function revealCard(card, shouldOpen = true) {
    if (!card) return;
    const group = card.closest('.cc-group');
    group?.classList.remove('is-collapsed','cc-group-filter-hidden','cc-group-search-hidden');
    group?.querySelector('.cc-group-head')?.setAttribute('aria-expanded','true');
    card.classList.remove('cc-card-search-hidden');
    if (shouldOpen) card.open = true;
    requestAnimationFrame(() => card.scrollIntoView({ behavior: 'smooth', block: 'start' }));
  }

  function renderRecents() {
    const container = state.host?.querySelector('.cc-recents');
    const list = state.host?.querySelector('.cc-recents-list');
    if (!container || !list) return;
    const recents = safeReadRecents().filter((item) => findCardByTitle(item.title));
    container.classList.toggle('has-items', recents.length > 0);
    list.innerHTML = '';
    recents.forEach((item) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'cc-recent-btn';
      button.textContent = item.title;
      button.title = item.title;
      button.addEventListener('click', () => revealCard(findCardByTitle(item.title)));
      list.appendChild(button);
    });
  }

  function setAllGroups(open) {
    state.host?.querySelectorAll('.cc-group').forEach((group) => {
      group.classList.toggle('is-collapsed', !open);
      group.querySelector('.cc-group-head')?.setAttribute('aria-expanded', String(open));
    });
    if (!open) state.cards.forEach((card) => { card.open = false; });
  }

  function applyGroupFilter(groupId) {
    state.activeGroup = groupId || 'all';
    state.host?.querySelectorAll('.cc-filter-chip').forEach((chip) => chip.classList.toggle('is-active', chip.dataset.ccFilter === state.activeGroup));
    state.host?.querySelectorAll('.cc-group').forEach((group) => group.classList.toggle('cc-group-filter-hidden', state.activeGroup !== 'all' && group.dataset.ccGroup !== state.activeGroup));
    const search = state.host?.querySelector('.cc-organizer-search');
    if (search?.value) filterCards(search.value);
  }

  function filterCards(query) {
    const term = normalize(query).trim();
    const words = term.split(/\s+/).filter(Boolean);
    let visibleTotal = 0;
    state.host?.querySelectorAll('.cc-group').forEach((groupNode) => {
      if (state.activeGroup !== 'all' && groupNode.dataset.ccGroup !== state.activeGroup) return;
      let visibleGroup = 0;
      groupNode.querySelectorAll('.cc-group-body>details').forEach((card) => {
        const haystack = cardSearchText(card);
        const matches = !words.length || words.every((word) => haystack.includes(word));
        card.classList.toggle('cc-card-search-hidden', !matches);
        if (matches) visibleGroup += 1;
        if (term && matches) card.open = true;
      });
      groupNode.classList.toggle('cc-group-search-hidden', visibleGroup === 0);
      if (term && visibleGroup) {
        groupNode.classList.remove('is-collapsed');
        groupNode.querySelector('.cc-group-head')?.setAttribute('aria-expanded','true');
      }
      visibleTotal += visibleGroup;
    });
    state.host?.querySelector('.cc-no-results')?.classList.toggle('is-visible', Boolean(term && visibleTotal === 0));
  }

  function bindCardUsage(card) {
    if (card.dataset.ccUsageBound === '1') return;
    card.dataset.ccUsageBound = '1';
    card.addEventListener('click', (event) => {
      const target = event.target instanceof Element ? event.target : null;
      if (!target?.closest('button,a,input,select,summary')) return;
      saveRecent(card);
    }, true);
  }

  function buildDashboard(host, cardCount) {
    const dashboard = document.createElement('section');
    dashboard.className = 'cc-dashboard';
    dashboard.innerHTML = `<div class="cc-dashboard-top"><div><h2 class="cc-dashboard-title">🛠️ Centro di controllo</h2><p class="cc-dashboard-subtitle">Trova subito ciò che ti serve. Le funzioni restano identiche: cambia soltanto l'organizzazione.</p></div><span class="cc-dashboard-count">${cardCount} funzioni</span></div><div class="cc-quick-actions"><button class="cc-quick-btn" type="button" data-cc-action="open">▾ Apri tutto</button><button class="cc-quick-btn" type="button" data-cc-action="close">▴ Chiudi tutto</button><button class="cc-quick-btn" type="button" data-cc-action="performance">⚡ Firestore</button></div>`;
    dashboard.querySelector('[data-cc-action="open"]')?.addEventListener('click', () => setAllGroups(true));
    dashboard.querySelector('[data-cc-action="close"]')?.addEventListener('click', () => setAllGroups(false));
    dashboard.querySelector('[data-cc-action="performance"]')?.addEventListener('click', () => {
      applyGroupFilter('performance');
      const performance = host.querySelector('[data-cc-group="performance"]');
      performance?.classList.remove('is-collapsed');
      performance?.querySelector('.cc-group-head')?.setAttribute('aria-expanded','true');
      performance?.scrollIntoView({ behavior:'smooth', block:'start' });
    });
    return dashboard;
  }

  function buildFilterStrip(host) {
    const strip = document.createElement('div');
    strip.className = 'cc-filter-strip';
    const all = [{ id:'all', icon:'✨', title:'Tutto' }, ...GROUPS];
    all.forEach((group) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = `cc-filter-chip${group.id === 'all' ? ' is-active' : ''}`;
      button.dataset.ccFilter = group.id;
      button.textContent = `${group.icon} ${group.title}`;
      button.addEventListener('click', () => applyGroupFilter(group.id));
      strip.appendChild(button);
    });
    host.appendChild(strip);
  }

  function organize() {
    const root = findRoot();
    if (!root) return false;
    if (root.dataset.ccOrganizerInstalled === '1') return true;
    const cards = collectCards(root);
    if (!cards.length) return false;
    ensureStyles();

    const anchorParent = cards[0].parentElement || root;
    const anchor = cards[0];
    const host = document.createElement('div');
    host.className = 'cc-organizer';
    host.appendChild(buildDashboard(host, cards.length));
    host.insertAdjacentHTML('beforeend', `<div class="cc-organizer-tools"><div class="cc-search-wrap"><span class="cc-search-icon" aria-hidden="true">🔎</span><input class="cc-organizer-search" type="search" autocomplete="off" placeholder="Cerca funzione, es. backup o utenti…" aria-label="Cerca funzione nel Centro di controllo"></div><button class="cc-organizer-clear" type="button" aria-label="Cancella ricerca">✕</button><p class="cc-organizer-hint">Puoi cercare anche più parole, ad esempio “letture Firestore”.</p></div><section class="cc-recents" aria-label="Ultime funzioni usate"><div class="cc-recents-head"><h3 class="cc-recents-title">🕘 Ultime usate</h3><button class="cc-recents-clear" type="button">Cancella</button></div><div class="cc-recents-list"></div></section><div class="cc-no-results"><strong>Nessuna funzione trovata</strong>Prova con una parola diversa o torna su “Tutto”.</div>`);
    buildFilterStrip(host);

    const groupNodes = new Map();
    GROUPS.forEach((group) => {
      const node = buildGroup(group);
      groupNodes.set(group.id, node);
      host.appendChild(node);
    });
    anchorParent.insertBefore(host, anchor);

    cards.forEach((card) => {
      const group = classify(card);
      card.dataset.ccOriginalTitle = cardTitle(card);
      card.dataset.ccGroup = group.id;
      card.dataset.ccGroupTitle = group.title;
      bindCardUsage(card);
      groupNodes.get(group.id)?.querySelector('.cc-group-body')?.appendChild(card);
    });

    groupNodes.forEach((node) => {
      const count = node.querySelectorAll('.cc-group-body>details').length;
      node.querySelector('.cc-group-count').textContent = String(count);
      if (!count) node.remove();
    });

    const search = host.querySelector('.cc-organizer-search');
    host.querySelector('.cc-organizer-clear')?.addEventListener('click', () => {
      if (search) search.value = '';
      applyGroupFilter('all');
      filterCards('');
      search?.focus();
    });
    search?.addEventListener('input', () => filterCards(search.value));
    host.querySelector('.cc-recents-clear')?.addEventListener('click', () => {
      try { localStorage.removeItem(RECENTS_KEY); } catch (_) {}
      renderRecents();
    });

    root.dataset.ccOrganizerInstalled = '1';
    root.classList.add('control-center-organized');
    state.root = root;
    state.host = host;
    state.cards = cards;
    renderRecents();
    return true;
  }

  function install() {
    if (organize()) return;
    if (state.observer) return;
    const scope = findRoot() || document.body;
    state.observer = new MutationObserver(() => {
      if (!organize()) return;
      state.observer?.disconnect();
      state.observer = null;
    });
    state.observer.observe(scope, { childList:true, subtree:true });
  }

  document.getElementById('open-control-center-btn')?.addEventListener('click', () => setTimeout(install, 0), true);
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once:true });
  else install();

  window.HeraControlCenterOrganizer = {
    installed: true,
    version: '2.0.0',
    refresh: organize,
    openGroup: applyGroupFilter,
    getState: () => ({ organized:Boolean(state.root?.dataset.ccOrganizerInstalled === '1'), cards:state.cards.length, activeGroup:state.activeGroup })
  };
})();

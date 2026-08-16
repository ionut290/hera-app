(() => {
  'use strict';

  if (window.HeraControlCenterOrganizer?.installed) return;

  const FIRESTORE_CARD_ID = 'firestore-usage-control-card';
  const RECENT_KEY = 'hera_control_center_recent_v2';
  const GROUPS = [
    { id: 'operations', icon: '🧭', title: 'Operatività', description: 'Commesse, impianti, squadre e ore', tone: 'blue', keywords: ['commessa','commesse','impianto','impianti','squadra','squadre','ore','programmazione','operativ'] },
    { id: 'people', icon: '👥', title: 'Persone e accessi', description: 'Utenti, personale, account e mezzi', tone: 'green', keywords: ['utente','utenti','personale','accesso','accessi','account','password','profilo','mezzi','mezzo'] },
    { id: 'communication', icon: '🔔', title: 'Comunicazioni', description: 'Notifiche, banner, avvisi e segnalazioni', tone: 'orange', keywords: ['notific','banner','avvisi','segnalaz','comunicaz','informazioni utili'] },
    { id: 'documents', icon: '📁', title: 'Documenti e dati', description: 'Backup, PDF, POS, Drive e archivi', tone: 'violet', keywords: ['document','pdf','pos','drive','backup','archivio','esporta','importa'] },
    { id: 'performance', icon: '⚡', title: 'Prestazioni', description: 'Diagnostica, letture, cache e velocità', tone: 'cyan', keywords: ['firestore','letture','scritture','listener','prestaz','performance','diagnostic','cache','storage','quota','veloc','consumo'] },
    { id: 'security', icon: '🛡️', title: 'Sicurezza', description: 'Protezione, manutenzione, errori e reset', tone: 'red', keywords: ['sicurezza','protezione','ripristino','manutenz','pulizia','reset','errore','log'] },
    { id: 'other', icon: '🧰', title: 'Altri strumenti', description: 'Utility e funzioni amministrative', tone: 'slate', keywords: [] }
  ];

  const state = { root: null, dashboard: null, workspace: null, cards: [], observer: null, activeGroup: '' };
  const normalize = (value) => String(value || '').toLocaleLowerCase('it-IT').normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  function findRoot() { return document.getElementById('control-center-content') || document.getElementById('control-center-page'); }

  function cardTitle(card) {
    const summary = card?.querySelector(':scope > summary');
    const heading = card?.querySelector(':scope > h2, :scope > h3, :scope > header h2, :scope > header h3, .fs-usage-title');
    return String(summary?.textContent || heading?.textContent || card?.getAttribute('aria-label') || card?.dataset?.ccOriginalTitle || 'Strumento').trim();
  }

  function cardText(card) { return normalize(`${cardTitle(card)} ${card?.textContent || ''}`); }

  function isToolCard(element, root) {
    if (!element || element === root || element.id === FIRESTORE_CARD_ID) return false;
    if (element.closest('.cc-dashboard-v2, .cc-workspace-v2')) return false;
    if (element.closest('details') && element.tagName !== 'DETAILS') return false;
    return element.matches('details, .control-card, .control-center-card');
  }

  function collectCards(root) {
    const candidates = Array.from(root.querySelectorAll('details, .control-card, .control-center-card'));
    return candidates.filter((card) => {
      if (!isToolCard(card, root)) return false;
      return !candidates.some((other) => other !== card && other.contains(card) && isToolCard(other, root));
    });
  }

  function classify(card) {
    const haystack = cardText(card);
    let best = GROUPS[GROUPS.length - 1];
    let bestScore = 0;
    GROUPS.slice(0, -1).forEach((group) => {
      const score = group.keywords.reduce((sum, keyword) => sum + (haystack.includes(normalize(keyword)) ? 1 : 0), 0);
      if (score > bestScore) { bestScore = score; best = group; }
    });
    return best;
  }

  function readRecent() { try { return JSON.parse(localStorage.getItem(RECENT_KEY) || '[]').filter(Boolean).slice(0, 4); } catch (_) { return []; } }
  function saveRecent(title) { if (!title) return; try { localStorage.setItem(RECENT_KEY, JSON.stringify([title, ...readRecent().filter((item) => item !== title)].slice(0, 4))); } catch (_) {} }

  function ensureStyles() {
    if (document.getElementById('control-center-organizer-style-v3')) return;
    const style = document.createElement('style');
    style.id = 'control-center-organizer-style-v3';
    style.textContent = `
      .cc-dashboard-v2{display:grid;gap:10px;margin:8px 0 14px}.cc-dashboard-v2.is-hidden{display:none!important}
      .cc-dashboard-head{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:4px 2px}.cc-dashboard-head h2{margin:0;font-size:1.05rem}.cc-dashboard-head small{color:#64748b}
      .cc-searchbar{display:grid;grid-template-columns:minmax(0,1fr) 44px;gap:7px}.cc-searchbar input{width:100%;height:44px;border:1px solid #cbd8e5;border-radius:13px;background:#f8fbff;padding:0 13px;font-size:16px}.cc-searchbar button{height:44px;border:1px solid #d4dde7;border-radius:13px;background:#f8fafc;font-weight:900}
      .cc-widget-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}.cc-widget{position:relative;display:flex;flex-direction:column;align-items:flex-start;justify-content:space-between;gap:8px;min-height:112px;padding:13px;border:1px solid rgba(148,163,184,.25);border-radius:17px;text-align:left;box-shadow:0 5px 16px rgba(15,23,42,.055);overflow:hidden}.cc-widget::after{content:'';position:absolute;width:70px;height:70px;border-radius:50%;right:-26px;bottom:-30px;background:rgba(255,255,255,.32)}.cc-widget-icon{font-size:1.45rem}.cc-widget-title{display:block;font-size:.9rem;font-weight:900;color:#16263d}.cc-widget-desc{display:block;font-size:.69rem;line-height:1.18;color:#526275}.cc-widget-count{position:absolute;top:10px;right:10px;display:grid;place-items:center;min-width:25px;height:25px;padding:0 6px;border-radius:999px;background:rgba(255,255,255,.72);font-size:.7rem;font-weight:900;color:#334155}
      .cc-widget[data-tone='blue']{background:linear-gradient(145deg,#dfeeff,#edf5ff)}.cc-widget[data-tone='green']{background:linear-gradient(145deg,#def6e9,#edf9f2)}.cc-widget[data-tone='orange']{background:linear-gradient(145deg,#fff0dc,#fff7ea)}.cc-widget[data-tone='violet']{background:linear-gradient(145deg,#eee8ff,#f6f2ff)}.cc-widget[data-tone='cyan']{background:linear-gradient(145deg,#dff5fb,#ecf9fc)}.cc-widget[data-tone='red']{background:linear-gradient(145deg,#ffe7e7,#fff2f2)}.cc-widget[data-tone='slate']{background:linear-gradient(145deg,#e9eef4,#f4f7fa)}
      .cc-recent{display:none;gap:6px;overflow-x:auto;padding:1px 0 3px}.cc-recent.has-items{display:flex}.cc-recent-btn{flex:0 0 auto;border:1px solid #d8e2ec;border-radius:999px;background:#f8fafc;padding:7px 10px;font-size:.72rem;font-weight:800;color:#3b4a5d}
      .cc-workspace-v2{display:none;margin:8px 0 16px;border:1px solid #d3dfeb;border-radius:18px;background:linear-gradient(180deg,#f3f7fb,#eef3f7);overflow:hidden}.cc-workspace-v2.is-open{display:block}.cc-workspace-head{position:sticky;top:0;z-index:5;display:flex;align-items:center;gap:8px;padding:9px;background:rgba(238,244,249,.96);backdrop-filter:blur(12px);border-bottom:1px solid #d8e2eb}.cc-workspace-back{width:40px;height:40px;border:1px solid #ccd8e3;border-radius:12px;background:#fff;font-size:1rem;font-weight:900}.cc-workspace-heading{flex:1;min-width:0}.cc-workspace-heading strong{display:block}.cc-workspace-heading small{display:block;color:#64748b;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.cc-workspace-body{display:grid;gap:6px;padding:7px}.cc-workspace-body>details,.cc-workspace-body>.control-card,.cc-workspace-body>.control-center-card{margin:0!important;border-radius:13px!important;box-shadow:none!important;background:#fff!important;border:1px solid #dbe4ed!important}.cc-workspace-body>details>summary{min-height:44px;padding:10px!important;font-weight:800}.cc-source-hidden{display:none!important}
      #${FIRESTORE_CARD_ID}{display:block!important;margin:8px 0 12px!important;border-radius:17px!important;box-shadow:0 6px 18px rgba(15,23,42,.06)!important}
      @media(min-width:760px){.cc-widget-grid{grid-template-columns:repeat(4,minmax(0,1fr))}.cc-widget{min-height:125px}.cc-workspace-body{grid-template-columns:repeat(2,minmax(0,1fr));align-items:start}}
      @media(max-width:420px){.cc-dashboard-v2{gap:7px}.cc-widget-grid{gap:6px}.cc-widget{min-height:99px;padding:10px;border-radius:14px}.cc-widget-icon{font-size:1.25rem}.cc-widget-title{font-size:.82rem}.cc-widget-desc{font-size:.64rem}.cc-workspace-v2{border-radius:15px}.cc-workspace-body{padding:5px;gap:5px}}
    `;
    document.head.appendChild(style);
  }

  function groupCards(groupId) { return state.cards.filter((card) => card.dataset.ccGroup === groupId); }

  function restoreAllToWorkspace() {
    if (!state.workspace) return;
    const body = state.workspace.querySelector('.cc-workspace-body');
    state.cards.forEach((card) => {
      card.classList.add('cc-source-hidden');
      if (card.parentElement !== body) body.appendChild(card);
    });
  }

  function renderRecent() {
    const holder = state.dashboard?.querySelector('.cc-recent');
    if (!holder) return;
    const available = new Set(state.cards.map(cardTitle));
    const recent = readRecent().filter((title) => available.has(title));
    holder.innerHTML = '';
    holder.classList.toggle('has-items', recent.length > 0);
    recent.forEach((title) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'cc-recent-btn';
      button.textContent = `↗ ${title}`;
      button.addEventListener('click', () => openCardByTitle(title));
      holder.appendChild(button);
    });
  }

  function recordCardUse(card) { const title = cardTitle(card); saveRecent(title); renderRecent(); }

  function showCards(cards, title, description) {
    if (!state.workspace) return;
    restoreAllToWorkspace();
    let visible = 0;
    state.cards.forEach((card) => {
      const show = cards.includes(card);
      card.classList.toggle('cc-source-hidden', !show);
      if (show) visible += 1;
    });
    state.workspace.querySelector('[data-cc-workspace-title]').textContent = title || 'Strumenti';
    state.workspace.querySelector('[data-cc-workspace-subtitle]').textContent = description || `${visible} funzioni`;
    state.dashboard?.classList.add('is-hidden');
    state.workspace.classList.add('is-open');
    state.workspace.scrollIntoView({ block: 'start', behavior: 'auto' });
  }

  function openGroup(group) { state.activeGroup = group.id; showCards(groupCards(group.id), `${group.icon} ${group.title}`, group.description); }

  function closeWorkspace() {
    state.workspace?.classList.remove('is-open');
    state.dashboard?.classList.remove('is-hidden');
    restoreAllToWorkspace();
    state.activeGroup = '';
    state.dashboard?.scrollIntoView({ block: 'start', behavior: 'auto' });
  }

  function openCardByTitle(title) {
    const card = state.cards.find((item) => cardTitle(item) === title);
    if (!card) return;
    showCards([card], cardTitle(card), 'Apertura rapida');
    card.classList.remove('cc-source-hidden');
    if (card.tagName === 'DETAILS') card.open = true;
    recordCardUse(card);
  }

  function search(query) {
    const words = normalize(query).trim().split(/\s+/).filter(Boolean);
    if (!words.length) { closeWorkspace(); return; }
    const matches = state.cards.filter((card) => { const text = cardText(card); return words.every((word) => text.includes(word)); });
    showCards(matches, '🔎 Risultati ricerca', matches.length ? `${matches.length} funzioni trovate` : 'Nessuna funzione trovata');
  }

  function createDashboard() {
    const dashboard = document.createElement('section');
    dashboard.className = 'cc-dashboard-v2';
    dashboard.innerHTML = `<div class="cc-dashboard-head"><div><h2>Centro di controllo</h2><small>Scegli un widget</small></div></div><div class="cc-searchbar"><input type="search" autocomplete="off" placeholder="🔎 Cerca funzione…" aria-label="Cerca nel Centro di controllo"><button type="button" aria-label="Cancella ricerca">✕</button></div><div class="cc-recent" aria-label="Ultime funzioni usate"></div><div class="cc-widget-grid"></div>`;
    const grid = dashboard.querySelector('.cc-widget-grid');
    GROUPS.forEach((group) => {
      const count = groupCards(group.id).length;
      if (!count) return;
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'cc-widget';
      button.dataset.tone = group.tone;
      button.innerHTML = `<span class="cc-widget-count">${count}</span><span class="cc-widget-icon" aria-hidden="true">${group.icon}</span><span><span class="cc-widget-title">${group.title}</span><span class="cc-widget-desc">${group.description}</span></span>`;
      button.addEventListener('click', () => openGroup(group));
      grid.appendChild(button);
    });
    const input = dashboard.querySelector('input');
    let timer = 0;
    input.addEventListener('input', () => { clearTimeout(timer); timer = window.setTimeout(() => search(input.value), 80); });
    dashboard.querySelector('.cc-searchbar button').addEventListener('click', () => { input.value = ''; closeWorkspace(); input.focus(); });
    return dashboard;
  }

  function createWorkspace() {
    const workspace = document.createElement('section');
    workspace.className = 'cc-workspace-v2';
    workspace.innerHTML = `<header class="cc-workspace-head"><button class="cc-workspace-back" type="button" aria-label="Torna ai widget">←</button><div class="cc-workspace-heading"><strong data-cc-workspace-title>Strumenti</strong><small data-cc-workspace-subtitle></small></div></header><div class="cc-workspace-body"></div>`;
    workspace.querySelector('.cc-workspace-back').addEventListener('click', closeWorkspace);
    workspace.addEventListener('click', (event) => { const card = event.target.closest('details, .control-card, .control-center-card'); if (card && state.cards.includes(card)) recordCardUse(card); }, true);
    return workspace;
  }

  function positionDashboard(root) {
    if (!state.dashboard || !state.workspace) return;
    const graph = document.getElementById(FIRESTORE_CARD_ID);
    if (graph?.parentElement === root) {
      graph.insertAdjacentElement('afterend', state.dashboard);
      state.dashboard.insertAdjacentElement('afterend', state.workspace);
    } else {
      root.prepend(state.workspace);
      root.prepend(state.dashboard);
    }
  }

  function absorbDynamicCards(root) {
    const fresh = collectCards(root).filter((card) => !state.cards.includes(card));
    if (!fresh.length) return;
    fresh.forEach((card) => {
      card.dataset.ccOriginalIndex = String(state.cards.length);
      card.dataset.ccOriginalTitle = cardTitle(card);
      card.dataset.ccGroup = classify(card).id;
      state.cards.push(card);
      card.classList.add('cc-source-hidden');
      state.workspace?.querySelector('.cc-workspace-body')?.appendChild(card);
    });
    const oldDashboard = state.dashboard;
    const replacement = createDashboard();
    oldDashboard?.replaceWith(replacement);
    state.dashboard = replacement;
    renderRecent();
  }

  function organize() {
    const root = findRoot();
    if (!root) return false;
    if (root.dataset.ccOrganizerInstalledV3 === '1') { positionDashboard(root); return true; }

    const cards = collectCards(root);
    if (!cards.length) return false;
    ensureStyles();
    state.root = root;
    state.cards = cards;
    cards.forEach((card, index) => { card.dataset.ccOriginalIndex = String(index); card.dataset.ccOriginalTitle = cardTitle(card); card.dataset.ccGroup = classify(card).id; });
    state.dashboard = createDashboard();
    state.workspace = createWorkspace();
    positionDashboard(root);
    restoreAllToWorkspace();
    renderRecent();
    root.dataset.ccOrganizerInstalledV3 = '1';
    root.classList.add('control-center-organized-v3');

    const observer = new MutationObserver(() => {
      absorbDynamicCards(root);
      const graph = document.getElementById(FIRESTORE_CARD_ID);
      if (graph && state.dashboard?.previousElementSibling !== graph) positionDashboard(root);
    });
    observer.observe(root, { childList: true });
    state.observer = observer;
    return true;
  }

  function install() {
    if (organize()) return;
    const scope = findRoot() || document.body;
    if (state.observer) return;
    state.observer = new MutationObserver(() => { if (!organize()) return; state.observer?.disconnect(); state.observer = null; });
    state.observer.observe(scope, { childList: true, subtree: true });
  }

  document.getElementById('open-control-center-btn')?.addEventListener('click', () => window.setTimeout(install, 0), true);
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true }); else install();

  window.HeraControlCenterOrganizer = {
    installed: true,
    version: '3.0.0',
    refresh: organize,
    getState: () => ({ organized: Boolean(state.root?.dataset.ccOrganizerInstalledV3 === '1'), cards: state.cards.length, activeGroup: state.activeGroup })
  };
})();
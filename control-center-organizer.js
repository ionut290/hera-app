(() => {
  'use strict';

  if (window.HeraControlCenterOrganizer?.installed) return;

  const GROUPS = [
    { id: 'operations', icon: '🧭', title: 'Operatività', keywords: ['commessa', 'commesse', 'impianto', 'impianti', 'squadra', 'squadre', 'ore', 'programmazione', 'operativ'] },
    { id: 'people', icon: '👥', title: 'Persone e accessi', keywords: ['utente', 'utenti', 'personale', 'accesso', 'accessi', 'account', 'password', 'profilo', 'mezzi', 'mezzo'] },
    { id: 'communication', icon: '🔔', title: 'Comunicazioni', keywords: ['notific', 'banner', 'avvisi', 'segnalaz', 'comunicaz', 'informazioni utili'] },
    { id: 'documents', icon: '📁', title: 'Documenti e dati', keywords: ['document', 'pdf', 'pos', 'drive', 'backup', 'archivio', 'esporta', 'importa'] },
    { id: 'performance', icon: '⚡', title: 'Firestore e prestazioni', keywords: ['firestore', 'letture', 'scritture', 'listener', 'prestaz', 'performance', 'diagnostic', 'cache', 'storage', 'quota', 'veloc', 'consumo'] },
    { id: 'security', icon: '🛡️', title: 'Sicurezza e manutenzione', keywords: ['sicurezza', 'protezione', 'ripristino', 'manutenz', 'pulizia', 'reset', 'errore', 'log'] },
    { id: 'other', icon: '🧰', title: 'Altri strumenti', keywords: [] }
  ];

  const state = { root: null, host: null, observer: null, cards: [] };

  function normalize(value) {
    return String(value || '')
      .toLocaleLowerCase('it-IT')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  }

  function findRoot() {
    return document.getElementById('control-center-content') || document.getElementById('control-center-page');
  }

  function cardTitle(card) {
    return String(card?.querySelector('summary')?.textContent || card?.getAttribute('aria-label') || '').trim();
  }

  function classify(card) {
    const haystack = normalize(`${cardTitle(card)} ${card?.textContent || ''}`);
    let best = GROUPS[GROUPS.length - 1];
    let score = 0;
    GROUPS.slice(0, -1).forEach((group) => {
      const current = group.keywords.reduce((total, keyword) => total + (haystack.includes(normalize(keyword)) ? 1 : 0), 0);
      if (current > score) {
        score = current;
        best = group;
      }
    });
    return best;
  }

  function ensureStyles() {
    if (document.getElementById('control-center-organizer-style')) return;
    const style = document.createElement('style');
    style.id = 'control-center-organizer-style';
    style.textContent = `
      .cc-organizer{display:grid;gap:12px;margin:8px 0 18px}
      .cc-organizer-tools{position:sticky;top:0;z-index:8;display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px;padding:10px;background:linear-gradient(135deg,#eaf3ff,#edf8f3);border:1px solid #d5e2ee;border-radius:16px;box-shadow:0 6px 18px rgba(30,64,100,.08)}
      .cc-organizer-search{min-height:44px;width:100%;padding:10px 13px;border:1px solid #bfd0df;border-radius:12px;background:#fff;font-size:16px}
      .cc-organizer-clear{min-width:44px;min-height:44px;border:1px solid #cbd8e5;border-radius:12px;background:#f8fbff;font-weight:800}
      .cc-organizer-hint{grid-column:1/-1;margin:0;color:#5c6b7a;font-size:.78rem}
      .cc-group{border:1px solid #d7e2ec;border-radius:17px;background:rgba(244,248,252,.9);overflow:hidden;box-shadow:0 5px 15px rgba(15,23,42,.045)}
      .cc-group-head{width:100%;display:flex;align-items:center;gap:10px;min-height:48px;padding:10px 13px;border:0;background:linear-gradient(90deg,#edf4fb,#f1f6f8);color:#1d2d44;text-align:left;font-weight:900}
      .cc-group-icon{font-size:1.15rem}.cc-group-title{flex:1}.cc-group-count{display:inline-grid;place-items:center;min-width:27px;height:27px;padding:0 7px;border-radius:999px;background:#dbe9f7;color:#315477;font-size:.75rem}.cc-group-chevron{font-size:1rem;transition:transform .16s ease}.cc-group.is-collapsed .cc-group-chevron{transform:rotate(-90deg)}
      .cc-group-body{display:grid;gap:7px;padding:7px}.cc-group.is-collapsed .cc-group-body{display:none}
      .cc-group-body>details{margin:0!important;border-radius:13px!important;box-shadow:none!important;border-color:#dce5ed!important;background:#fff!important}
      .cc-group-body>details>summary{min-height:44px;padding:10px 12px!important;font-weight:800}
      .cc-card-search-hidden{display:none!important}.cc-group-search-hidden{display:none!important}
      .cc-no-results{display:none;padding:18px;border:1px dashed #c8d4df;border-radius:14px;background:#f8fafc;text-align:center;color:#64748b}.cc-no-results.is-visible{display:block}
      @media(max-width:700px){.cc-organizer{gap:8px}.cc-organizer-tools{top:4px;padding:7px;border-radius:14px}.cc-organizer-search{min-height:42px}.cc-organizer-clear{min-height:42px}.cc-group{border-radius:14px}.cc-group-head{min-height:44px;padding:8px 10px}.cc-group-body{gap:4px;padding:4px}.cc-group-body>details>summary{min-height:42px;padding:8px 10px!important}}
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
    const initiallyOpen = group.id === 'performance';
    section.className = `cc-group${initiallyOpen ? '' : ' is-collapsed'}`;
    section.dataset.ccGroup = group.id;
    section.innerHTML = `<button class="cc-group-head" type="button" aria-expanded="${String(initiallyOpen)}"><span class="cc-group-icon" aria-hidden="true">${group.icon}</span><span class="cc-group-title">${group.title}</span><span class="cc-group-count">0</span><span class="cc-group-chevron" aria-hidden="true">⌄</span></button><div class="cc-group-body"></div>`;
    const button = section.querySelector('.cc-group-head');
    button.addEventListener('click', () => {
      const collapsed = section.classList.toggle('is-collapsed');
      button.setAttribute('aria-expanded', String(!collapsed));
    });
    return section;
  }

  function filterCards(query) {
    const term = normalize(query).trim();
    let visibleTotal = 0;
    state.host?.querySelectorAll('.cc-group').forEach((groupNode) => {
      let visibleGroup = 0;
      groupNode.querySelectorAll('.cc-group-body>details').forEach((card) => {
        const matches = !term || normalize(`${cardTitle(card)} ${card.textContent || ''}`).includes(term);
        card.classList.toggle('cc-card-search-hidden', !matches);
        if (matches) visibleGroup += 1;
        if (term && matches) card.open = true;
      });
      groupNode.classList.toggle('cc-group-search-hidden', visibleGroup === 0);
      if (term && visibleGroup) {
        groupNode.classList.remove('is-collapsed');
        groupNode.querySelector('.cc-group-head')?.setAttribute('aria-expanded', 'true');
      }
      visibleTotal += visibleGroup;
    });
    state.host?.querySelector('.cc-no-results')?.classList.toggle('is-visible', Boolean(term && visibleTotal === 0));
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
    host.innerHTML = `<div class="cc-organizer-tools"><input class="cc-organizer-search" type="search" autocomplete="off" placeholder="🔎 Cerca nel Centro di controllo…" aria-label="Cerca funzione nel Centro di controllo"><button class="cc-organizer-clear" type="button" aria-label="Cancella ricerca">✕</button><p class="cc-organizer-hint">Scrivi ad esempio: Firestore, backup, utenti, ore, notifiche…</p></div><div class="cc-no-results">Nessuna funzione trovata. Prova con un'altra parola.</div>`;

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
      groupNodes.get(group.id)?.querySelector('.cc-group-body')?.appendChild(card);
    });

    groupNodes.forEach((node) => {
      const count = node.querySelectorAll('.cc-group-body>details').length;
      node.querySelector('.cc-group-count').textContent = String(count);
      if (!count) node.remove();
    });

    const search = host.querySelector('.cc-organizer-search');
    const clear = host.querySelector('.cc-organizer-clear');
    search?.addEventListener('input', () => filterCards(search.value));
    clear?.addEventListener('click', () => {
      if (search) search.value = '';
      filterCards('');
      search?.focus();
    });

    root.dataset.ccOrganizerInstalled = '1';
    root.classList.add('control-center-organized');
    state.root = root;
    state.host = host;
    state.cards = cards;
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
    state.observer.observe(scope, { childList: true, subtree: true });
  }

  document.getElementById('open-control-center-btn')?.addEventListener('click', () => setTimeout(install, 0), true);
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true });
  else install();

  window.HeraControlCenterOrganizer = {
    installed: true,
    version: '1.0.0',
    refresh: organize,
    getState: () => ({ organized: Boolean(state.root?.dataset.ccOrganizerInstalled === '1'), cards: state.cards.length })
  };
})();

(() => {
  'use strict';

  const LIST_ID = 'squadre-lista';
  const ACTION_ID = 'today-commesse-action';
  const LABEL_ID = 'today-commesse-count';

  function clean(value) {
    return String(value || '').replace(/\s+/g, ' ').trim();
  }

  function normalizeName(value) {
    let text = clean(value);
    if (!text) return '';
    text = text
      .replace(/^commessa\s*[:\-–—]?\s*/i, '')
      .replace(/^apri\s+/i, '')
      .replace(/^(squadra|operatori|mezzi|impianti)\s*[:\-–—]?\s*/i, '')
      .trim();
    if (!text || /^(commessa|commesse|nessuna commessa|squadra|operatori|mezzi|impianti)$/i.test(text)) return '';
    return text;
  }

  function looksLikeCommessaName(value) {
    const text = normalizeName(value);
    if (!text || text.length < 2 || text.length > 80) return '';
    if (/^\d{1,2}[\/:.-]\d{1,2}/.test(text)) return '';
    if (/^(oggi|domani|auto|nessun mezzo|nessuna squadra|caricamento)/i.test(text)) return '';
    return text;
  }

  function candidateFromCard(card) {
    if (!card) return '';

    const datasetNames = [
      card.dataset?.commessaNome,
      card.dataset?.commessaName,
      card.dataset?.nomeCommessa,
      card.dataset?.commessa
    ];
    for (const value of datasetNames) {
      const name = looksLikeCommessaName(value);
      if (name) return name;
    }

    const explicitSelectors = [
      '[data-commessa-nome]', '[data-commessa-name]', '[data-nome-commessa]', '[data-commessa]',
      '.squadra-commessa-name', '.squadre-commessa-name', '.commessa-name', '.commessa-nome',
      '[class*="commessa"][class*="name"]', '[class*="commessa"][class*="nome"]',
      'h2', 'h3', 'h4', '.card-title', '.squadre-title', '.squadra-title'
    ];

    for (const selector of explicitSelectors) {
      const nodes = card.querySelectorAll(selector);
      for (const node of nodes) {
        const attrValue = node.dataset?.commessaNome || node.dataset?.commessaName || node.dataset?.nomeCommessa || node.dataset?.commessa;
        const name = looksLikeCommessaName(attrValue || node.textContent);
        if (name) return name;
      }
    }

    const rawText = String(card.innerText || card.textContent || '');
    const labelled = rawText.match(/commessa\s*[:\-–—]\s*([^|•\n]{2,80})/i);
    if (labelled?.[1]) {
      const name = looksLikeCommessaName(labelled[1]);
      if (name) return name;
    }

    const lines = rawText.split(/\n+/).map(clean).filter(Boolean);
    for (const line of lines) {
      const name = looksLikeCommessaName(line);
      if (!name) continue;
      if (/^(squadra|operatori|mezzi|impianti|data|ore|nessun|apri|naviga|fatto|whatsapp|whazzup)\b/i.test(name)) continue;
      return name;
    }

    return '';
  }

  function assignedCommessaNames() {
    const list = document.getElementById(LIST_ID);
    if (!list) return [];

    const cards = Array.from(list.children).filter((node) => node.nodeType === 1 && !node.classList.contains('hidden'));
    const names = [];
    const seen = new Set();

    cards.forEach((card) => {
      const name = candidateFromCard(card);
      if (!name) return;
      const key = name.toLocaleLowerCase('it-IT');
      if (seen.has(key)) return;
      seen.add(key);
      names.push(name);
    });

    return names;
  }

  function syncTodayCommessaLabel() {
    const action = document.getElementById(ACTION_ID);
    const label = document.getElementById(LABEL_ID);
    const button = document.getElementById('today-commesse-btn');
    if (!action || !label || !button) return;

    const names = assignedCommessaNames();

    if (names.length === 1) {
      action.textContent = `APRI ${names[0]}`;
      label.textContent = '';
      label.title = '';
      button.setAttribute('aria-label', `Apri ${names[0]}`);
      return;
    }

    if (names.length > 1) {
      action.textContent = `APRI ${names.length} COMMESSE`;
      label.textContent = '';
      label.title = '';
      button.setAttribute('aria-label', 'Apri le commesse assegnate oggi');
      return;
    }

    /* Mantiene il fallback esistente finché le squadre non hanno finito di renderizzare. */
    const fallback = clean(label.textContent);
    action.textContent = fallback && !/^commessa$/i.test(fallback) ? `APRI ${fallback}` : 'APRI COMMESSA';
    label.textContent = '';
  }

  function install() {
    syncTodayCommessaLabel();
    const list = document.getElementById(LIST_ID);
    if (!list || list.dataset.todayAdaptiveBound === '1') return;
    list.dataset.todayAdaptiveBound = '1';

    let timer = 0;
    const observer = new MutationObserver(() => {
      window.clearTimeout(timer);
      timer = window.setTimeout(syncTodayCommessaLabel, 40);
    });
    observer.observe(list, {
      childList: true,
      subtree: true,
      characterData: true,
      attributes: true,
      attributeFilter: ['class', 'data-commessa-nome', 'data-commessa-name', 'data-nome-commessa', 'data-commessa']
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', install, { once: true });
  } else {
    install();
  }
})();

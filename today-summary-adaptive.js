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
      .trim();
    if (!text || /^(commessa|commesse|nessuna commessa)$/i.test(text)) return '';
    return text;
  }

  function candidateFromCard(card) {
    if (!card) return '';

    const datasetNames = [
      card.dataset?.commessaNome,
      card.dataset?.commessaName,
      card.dataset?.nomeCommessa
    ];
    for (const value of datasetNames) {
      const name = normalizeName(value);
      if (name) return name;
    }

    const explicit = card.querySelector(
      '[data-commessa-nome], [data-commessa-name], [data-nome-commessa], .squadra-commessa-name, .squadre-commessa-name, .commessa-name, .commessa-nome'
    );
    if (explicit) {
      const attrValue = explicit.dataset?.commessaNome || explicit.dataset?.commessaName || explicit.dataset?.nomeCommessa;
      const name = normalizeName(attrValue || explicit.textContent);
      if (name) return name;
    }

    const text = clean(card.textContent);
    const labelled = text.match(/(?:^|\s)commessa\s*[:\-–—]\s*([^|•\n]{2,80})/i);
    if (labelled?.[1]) return normalizeName(labelled[1]);

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
    if (!action || !label) return;

    const names = assignedCommessaNames();
    action.textContent = 'APRI';

    if (names.length === 1) {
      label.textContent = names[0];
      label.title = `Apri ${names[0]}`;
      document.getElementById('today-commesse-btn')?.setAttribute('aria-label', `Apri ${names[0]}`);
      return;
    }

    if (names.length > 1) {
      label.textContent = `${names.length} COMMESSE`;
      label.title = 'Apri le commesse assegnate oggi';
      document.getElementById('today-commesse-btn')?.setAttribute('aria-label', 'Apri le commesse assegnate oggi');
    }
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
    observer.observe(list, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ['class', 'data-commessa-nome', 'data-commessa-name', 'data-nome-commessa'] });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', install, { once: true });
  } else {
    install();
  }
})();

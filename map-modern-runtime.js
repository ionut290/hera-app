(() => {
  'use strict';

  if (window.__heraModernImpiantiMapInstalled) return;
  window.__heraModernImpiantiMapInstalled = true;

  const INLINE_SEARCH_ID = 'map-number-search-form';
  const FULLSCREEN_FORM_ID = 'map-fullscreen-number-search-form';
  const FULLSCREEN_INPUT_ID = 'map-fullscreen-number-search-input';
  const FULLSCREEN_PAGE_ID = 'map-fullscreen-page';

  const normalize = (value) => String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();

  const firstValue = (item, keys) => {
    for (const key of keys) {
      const value = item?.[key];
      if (value !== undefined && value !== null && String(value).trim()) return String(value).trim();
    }
    return '';
  };

  function getCurrentImpianti() {
    try {
      if (typeof currentImpianti !== 'undefined' && Array.isArray(currentImpianti)) return currentImpianti;
    } catch (_) {}
    try {
      if (Array.isArray(window.currentImpianti)) return window.currentImpianti;
    } catch (_) {}
    try {
      if (typeof impiantiByCommessaId !== 'undefined' && impiantiByCommessaId?.get) {
        const id = typeof selectedCommessaId !== 'undefined' ? selectedCommessaId : window.selectedCommessaId;
        const items = impiantiByCommessaId.get(id);
        if (Array.isArray(items)) return items;
      }
    } catch (_) {}
    return [];
  }

  function plantName(item, index) {
    return firstValue(item, [
      'denominazioneImpianto', 'denominazione_impianto', 'Denominazione Impianto',
      'denominazione', 'nomeImpianto', 'nome', 'impianto', 'name', 'title'
    ]) || `Impianto ${index + 1}`;
  }

  function plantComune(item) {
    return firstValue(item, ['comune', 'Comune', 'ubicazione', 'localita', 'località', 'city']);
  }

  function plantAddress(item) {
    return firstValue(item, [
      'descrizioneVia', 'descrizione_via', 'Descrizione via', 'indirizzo', 'Indirizzo',
      'via', 'address'
    ]);
  }

  function plantCode(item) {
    return firstValue(item, [
      'idSap', 'idSAP', 'ID SAP', 'idsap', 'sap', 'codiceSap', 'codiceSAP',
      'codice', 'code', 'id'
    ]);
  }

  function plantNumber(item, index) {
    const raw = firstValue(item, [
      'numero', 'Numero', 'numeroImpianto', 'numero_impianto', 'progressivo',
      'n', 'markerNumber', 'marker_number'
    ]);
    const parsed = Number.parseInt(raw, 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : index + 1;
  }

  function searchableText(item, index) {
    return normalize([
      plantName(item, index), plantComune(item), plantAddress(item), plantCode(item),
      firstValue(item, ['tipologiaImpianto', 'tipologia_impianto', 'Tipologia impianto', 'tipologia', 'tipo']),
      String(plantNumber(item, index))
    ].filter(Boolean).join(' '));
  }

  function removeInlineSearch() {
    document.getElementById(INLINE_SEARCH_ID)?.remove();
  }

  function removeRedPinControl() {
    const shell = document.querySelector('.commessa-map-shell');
    if (!shell) return;
    shell.querySelectorAll('button, [role="button"]').forEach((node) => {
      if (node.id === 'map-inline-fullscreen-btn' || node.closest('.leaflet-control-zoom')) return;
      const signature = normalize([
        node.textContent,
        node.getAttribute('title'),
        node.getAttribute('aria-label'),
        node.className
      ].filter(Boolean).join(' '));
      const raw = `${node.textContent || ''} ${node.className || ''}`;
      const hasPinGlyph = /📍|🔴|pin/i.test(raw);
      const isLocationControl = /(posizione|localizz|centra.*mappa|mia posizione|geolocat|location)/.test(signature);
      if (hasPinGlyph || isLocationControl) node.remove();
    });
  }

  function installModernStyles() {
    if (document.getElementById('hera-modern-map-style')) return;
    const style = document.createElement('style');
    style.id = 'hera-modern-map-style';
    style.textContent = `
      .commessa-map-shell {
        overflow: hidden;
        border-radius: 24px !important;
        border: 1px solid rgba(15, 23, 42, .08);
        box-shadow: 0 12px 30px rgba(15, 23, 42, .10);
        background: #eef3f7;
      }
      .commessa-map-shell #map { border-radius: 24px !important; }
      .commessa-map-shell #${INLINE_SEARCH_ID} { display: none !important; }
      #${FULLSCREEN_PAGE_ID} #${FULLSCREEN_FORM_ID} {
        display: flex !important;
        position: relative;
        flex: 1 1 360px;
        min-width: min(420px, calc(100vw - 32px));
        max-width: 680px;
        margin: 0 auto;
        padding: 0;
        background: transparent;
        box-shadow: none;
        z-index: 1200;
      }
      #${FULLSCREEN_FORM_ID} .hera-map-search-shell { position: relative; width: 100%; }
      #${FULLSCREEN_FORM_ID} .hera-map-search-icon {
        position: absolute; left: 16px; top: 50%; transform: translateY(-50%);
        z-index: 2; pointer-events: none; font-size: 17px; line-height: 1; opacity: .72;
      }
      #${FULLSCREEN_INPUT_ID} {
        width: 100% !important; height: 52px !important;
        padding: 0 48px 0 46px !important;
        border: 1px solid rgba(15, 23, 42, .10) !important;
        border-radius: 18px !important; outline: none !important;
        background: rgba(255, 255, 255, .98) !important; color: #111827 !important;
        font-size: 16px !important; font-weight: 600 !important;
        box-shadow: 0 8px 24px rgba(15, 23, 42, .16) !important;
        -webkit-appearance: none; appearance: none;
      }
      #${FULLSCREEN_INPUT_ID}:focus {
        border-color: rgba(37, 99, 235, .42) !important;
        box-shadow: 0 9px 28px rgba(15, 23, 42, .17), 0 0 0 3px rgba(37, 99, 235, .10) !important;
      }
      #${FULLSCREEN_FORM_ID} > button[type="submit"] { display: none !important; }
      .hera-map-search-clear {
        position: absolute; right: 10px; top: 50%; transform: translateY(-50%);
        width: 34px; height: 34px; border: 0; border-radius: 50%; background: transparent;
        color: #64748b; font-size: 22px; line-height: 34px; cursor: pointer;
      }
      .hera-map-search-clear[hidden], .hera-map-search-results[hidden] { display: none !important; }
      .hera-map-search-results {
        position: absolute; left: 0; right: 0; top: calc(100% + 8px);
        max-height: min(55vh, 420px); overflow: auto; padding: 7px;
        border: 1px solid rgba(15, 23, 42, .10); border-radius: 18px;
        background: rgba(255, 255, 255, .99); box-shadow: 0 16px 36px rgba(15, 23, 42, .20);
        -webkit-overflow-scrolling: touch;
      }
      .hera-map-search-result {
        display: grid; grid-template-columns: 38px minmax(0, 1fr); gap: 10px;
        width: 100%; padding: 10px 11px; border: 0; border-radius: 13px;
        background: transparent; color: #111827; text-align: left; cursor: pointer;
      }
      .hera-map-search-result:active, .hera-map-search-result:hover, .hera-map-search-result.is-active { background: #f1f5f9; }
      .hera-map-search-result-icon {
        display: grid; place-items: center; width: 38px; height: 38px; border-radius: 50%;
        background: #eef6ff; font-size: 18px;
      }
      .hera-map-search-result-main { min-width: 0; }
      .hera-map-search-result-title {
        display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        font-weight: 800; font-size: 14px;
      }
      .hera-map-search-result-meta {
        display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
        margin-top: 2px; color: #64748b; font-size: 12px; font-weight: 600;
      }
      .leaflet-control-zoom, .leaflet-control-layers, .leaflet-bar {
        border: 0 !important; border-radius: 14px !important; overflow: hidden;
        box-shadow: 0 5px 18px rgba(15, 23, 42, .18) !important;
      }
      .leaflet-control-zoom a, .leaflet-bar a {
        width: 42px !important; height: 42px !important; line-height: 42px !important;
        border: 0 !important; background: rgba(255,255,255,.97) !important; color: #1f2937 !important;
      }
      .leaflet-control-zoom a + a, .leaflet-bar a + a { border-top: 1px solid rgba(15,23,42,.08) !important; }
      .leaflet-tile-pane { filter: saturate(.88) contrast(.98) brightness(1.02); }
      .leaflet-marker-pane, .leaflet-tooltip-pane, .leaflet-popup-pane { filter: none !important; }
      .leaflet-popup-content-wrapper { border-radius: 18px !important; box-shadow: 0 12px 30px rgba(15,23,42,.18) !important; }
      .leaflet-control-attribution {
        padding: 2px 7px !important; border-radius: 8px 0 0 0 !important;
        background: rgba(255,255,255,.82) !important; backdrop-filter: blur(8px); font-size: 10px !important;
      }
      #map-fullscreen-view { background: #e9eff4; }
      @media (max-width: 720px) {
        #${FULLSCREEN_PAGE_ID} .map-fullscreen-toolbar { gap: 8px; }
        #${FULLSCREEN_PAGE_ID} #${FULLSCREEN_FORM_ID} { order: 20; flex-basis: 100%; min-width: 100%; max-width: none; }
      }
    `;
    document.head.appendChild(style);
  }

  function buildResultButton(item, index) {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'hera-map-search-result';
    button.dataset.plantIndex = String(index);

    const icon = document.createElement('span');
    icon.className = 'hera-map-search-result-icon';
    icon.setAttribute('aria-hidden', 'true');
    icon.textContent = '📍';

    const main = document.createElement('span');
    main.className = 'hera-map-search-result-main';
    const title = document.createElement('span');
    title.className = 'hera-map-search-result-title';
    title.textContent = plantName(item, index);
    const meta = document.createElement('span');
    meta.className = 'hera-map-search-result-meta';
    const pieces = [plantComune(item), plantAddress(item), plantCode(item) ? `ID SAP ${plantCode(item)}` : ''].filter(Boolean);
    meta.textContent = pieces.join(' • ') || `Impianto n. ${plantNumber(item, index)}`;
    main.append(title, meta);
    button.append(icon, main);
    return button;
  }

  function setupFullscreenSearch() {
    const form = document.getElementById(FULLSCREEN_FORM_ID);
    const input = document.getElementById(FULLSCREEN_INPUT_ID);
    if (!form || !input || form.dataset.modernPlantSearch === '1') return;
    form.dataset.modernPlantSearch = '1';

    input.type = 'search';
    input.removeAttribute('min');
    input.removeAttribute('max');
    input.removeAttribute('step');
    input.inputMode = 'search';
    input.placeholder = 'Cerca impianto, comune, indirizzo o ID SAP';
    input.setAttribute('aria-label', 'Cerca un impianto sulla mappa');
    input.setAttribute('autocomplete', 'off');
    input.setAttribute('spellcheck', 'false');

    const shell = document.createElement('div');
    shell.className = 'hera-map-search-shell';
    input.parentNode.insertBefore(shell, input);
    shell.appendChild(input);

    const searchIcon = document.createElement('span');
    searchIcon.className = 'hera-map-search-icon';
    searchIcon.setAttribute('aria-hidden', 'true');
    searchIcon.textContent = '🔎';
    shell.appendChild(searchIcon);

    const clear = document.createElement('button');
    clear.type = 'button';
    clear.className = 'hera-map-search-clear';
    clear.setAttribute('aria-label', 'Cancella ricerca');
    clear.textContent = '×';
    clear.hidden = true;
    shell.appendChild(clear);

    const results = document.createElement('div');
    results.className = 'hera-map-search-results';
    results.setAttribute('role', 'listbox');
    results.hidden = true;
    shell.appendChild(results);

    let matches = [];
    let activeIndex = -1;
    let delegatingNumericSearch = false;

    const hideResults = () => { results.hidden = true; activeIndex = -1; };
    const updateActiveResult = () => {
      const buttons = Array.from(results.querySelectorAll('.hera-map-search-result'));
      buttons.forEach((button, idx) => button.classList.toggle('is-active', idx === activeIndex));
      buttons[activeIndex]?.scrollIntoView?.({ block: 'nearest' });
    };

    const renderMatches = () => {
      const query = normalize(input.value);
      clear.hidden = !input.value;
      results.textContent = '';
      activeIndex = -1;
      if (!query) { matches = []; hideResults(); return; }

      matches = getCurrentImpianti()
        .map((item, index) => ({ item, index, text: searchableText(item, index) }))
        .filter(({ text }) => text.includes(query))
        .sort((a, b) => {
          const aName = normalize(plantName(a.item, a.index));
          const bName = normalize(plantName(b.item, b.index));
          return (aName.startsWith(query) ? 0 : 1) - (bName.startsWith(query) ? 0 : 1) || aName.localeCompare(bName, 'it');
        })
        .slice(0, 8);

      if (!matches.length) {
        const empty = document.createElement('div');
        empty.className = 'hera-map-search-result';
        empty.setAttribute('aria-disabled', 'true');
        empty.innerHTML = '<span class="hera-map-search-result-icon" aria-hidden="true">🔎</span><span class="hera-map-search-result-main"><span class="hera-map-search-result-title">Nessun impianto trovato</span><span class="hera-map-search-result-meta">Prova con nome, comune, indirizzo o ID SAP</span></span>';
        results.appendChild(empty);
        results.hidden = false;
        return;
      }

      matches.forEach(({ item, index }) => results.appendChild(buildResultButton(item, index)));
      results.hidden = false;
    };

    const delegateToExistingNumericSearch = (item, index) => {
      input.value = String(plantNumber(item, index));
      clear.hidden = false;
      hideResults();
      delegatingNumericSearch = true;
      try {
        if (typeof form.requestSubmit === 'function') form.requestSubmit();
        else form.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
      } finally {
        setTimeout(() => { delegatingNumericSearch = false; }, 0);
      }
    };

    const selectMatch = (match) => { if (match) delegateToExistingNumericSearch(match.item, match.index); };

    input.addEventListener('input', renderMatches);
    input.addEventListener('focus', () => { if (input.value) renderMatches(); });
    input.addEventListener('keydown', (event) => {
      if (results.hidden || !matches.length) return;
      if (event.key === 'ArrowDown') {
        event.preventDefault(); activeIndex = Math.min(activeIndex + 1, matches.length - 1); updateActiveResult();
      } else if (event.key === 'ArrowUp') {
        event.preventDefault(); activeIndex = Math.max(activeIndex - 1, 0); updateActiveResult();
      } else if (event.key === 'Enter' && activeIndex >= 0) {
        event.preventDefault(); event.stopPropagation(); selectMatch(matches[activeIndex]);
      } else if (event.key === 'Escape') hideResults();
    });

    results.addEventListener('pointerdown', (event) => {
      const button = event.target.closest('.hera-map-search-result[data-plant-index]');
      if (!button) return;
      event.preventDefault();
      const index = Number(button.dataset.plantIndex);
      selectMatch(matches.find((candidate) => candidate.index === index));
    });

    clear.addEventListener('click', () => {
      input.value = ''; clear.hidden = true; hideResults(); input.focus();
    });

    form.addEventListener('submit', (event) => {
      if (delegatingNumericSearch) return;
      const raw = input.value.trim();
      if (/^\d+$/.test(raw)) return;
      event.preventDefault();
      event.stopImmediatePropagation();
      const query = normalize(raw);
      if (!query) return;
      const match = getCurrentImpianti()
        .map((item, index) => ({ item, index, text: searchableText(item, index) }))
        .find(({ text }) => text.includes(query));
      if (match) selectMatch(match);
      else renderMatches();
    }, true);

    document.addEventListener('pointerdown', (event) => { if (!form.contains(event.target)) hideResults(); });
  }

  function refresh() {
    removeInlineSearch();
    removeRedPinControl();
    installModernStyles();
    setupFullscreenSearch();
  }

  function init() {
    refresh();
    const observer = new MutationObserver(refresh);
    observer.observe(document.documentElement, { childList: true, subtree: true });
    window.addEventListener('hashchange', refresh);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();

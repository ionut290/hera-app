(() => {
  'use strict';

  if (window.__heraMapSearchFocusV3Installed) return;
  window.__heraMapSearchFocusV3Installed = true;

  const INPUT_ID = 'map-fullscreen-number-search-input';
  const PAGE_ID = 'map-fullscreen-page';
  const VIEW_ID = 'map-fullscreen-view';
  const SUGGESTIONS_ID = 'hera-map-live-suggestions';
  const FLASH_CLASS = 'hera-search-match-flash';
  const SELECTED_CLASS = 'hera-search-selected-marker';
  const MAX_RESULTS = 8;
  const MAX_AUTO_CENTER = 6;

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
    try { if (typeof currentImpianti !== 'undefined' && Array.isArray(currentImpianti)) return currentImpianti; } catch (_) {}
    try { if (Array.isArray(window.currentImpianti)) return window.currentImpianti; } catch (_) {}
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
    return firstValue(item, ['denominazioneImpianto','denominazione_impianto','Denominazione Impianto','denominazione','nomeImpianto','nome','impianto','name','title']) || `Impianto ${index + 1}`;
  }
  function plantComune(item) { return firstValue(item, ['comune','Comune','ubicazione','localita','località','city']); }
  function plantAddress(item) { return firstValue(item, ['descrizioneVia','descrizione_via','Descrizione via','indirizzo','Indirizzo','via','address']); }
  function plantCode(item) { return firstValue(item, ['idSap','idSAP','ID SAP','idsap','sap','codiceSap','codiceSAP','codice','code','id']); }
  function plantNumber(item, index) {
    const raw = firstValue(item, ['numero','Numero','numeroImpianto','numero_impianto','progressivo','n','markerNumber','marker_number']);
    const parsed = Number.parseInt(raw, 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : index + 1;
  }

  function plantCoordinates(item) {
    const latRaw = firstValue(item, ['lat','latitude','Latitudine','latitudine','GPS(Y)','gpsY','gps_y','coordinateY','coordinate_y','y']);
    const lngRaw = firstValue(item, ['lng','lon','longitude','Longitudine','longitudine','GPS(X)','gpsX','gps_x','coordinateX','coordinate_x','x']);
    const lat = Number(String(latRaw).replace(',', '.'));
    const lng = Number(String(lngRaw).replace(',', '.'));
    if (Number.isFinite(lat) && Number.isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180) return [lat, lng];
    const combined = firstValue(item, ['coordinate','Coordinate','coordinateGPS','Coordinate GPS','gps','GPS']);
    const nums = combined ? (String(combined).match(/-?\d+(?:[.,]\d+)?/g)?.map((v) => Number(v.replace(',', '.'))) || []) : [];
    if (nums.length >= 2) {
      const [a,b] = nums;
      if (Math.abs(a) <= 90 && Math.abs(b) <= 180) return [a,b];
      if (Math.abs(b) <= 90 && Math.abs(a) <= 180) return [b,a];
    }
    return null;
  }

  function matchScore(item, index, query) {
    const name = normalize(plantName(item,index));
    const comune = normalize(plantComune(item));
    const address = normalize(plantAddress(item));
    const code = normalize(plantCode(item));
    const number = normalize(String(plantNumber(item,index)));
    const values = [name, comune, address, code, number].filter(Boolean);
    if (!values.length) return Infinity;
    if (number === query || code === query || name === query) return 0;
    if (values.some((v) => v.startsWith(query))) return 1;
    if (values.some((v) => v.split(/\s+/).some((word) => word.startsWith(query)))) return 2;
    if (values.some((v) => v.includes(query))) return 3;
    return Infinity;
  }

  function findMatches(query) {
    return getCurrentImpianti()
      .map((item,index) => ({ item, index, score: matchScore(item,index,query) }))
      .filter((entry) => Number.isFinite(entry.score))
      .sort((a,b) => a.score - b.score || plantNumber(a.item,a.index) - plantNumber(b.item,b.index));
  }

  function getFullscreenMap() {
    const candidates = [window.mapFullscreen,window.fullscreenMap,window.mapFullScreen,window.fullScreenMap,window.mapFull,window.map].filter(Boolean);
    for (const candidate of candidates) {
      try {
        if (typeof candidate?.fitBounds !== 'function' || typeof candidate?.setView !== 'function') continue;
        const container = candidate.getContainer?.();
        if (container?.id === VIEW_ID || container?.closest?.(`#${PAGE_ID}`)) return candidate;
      } catch (_) {}
    }
    return null;
  }

  function clearFlashes() {
    document.querySelectorAll(`.${FLASH_CLASS}`).forEach((node) => node.classList.remove(FLASH_CLASS));
  }
  function clearSelected() {
    document.querySelectorAll(`.${SELECTED_CLASS}`).forEach((node) => node.classList.remove(SELECTED_CLASS));
  }

  function markerCandidatesForNumber(number) {
    const view = document.getElementById(VIEW_ID);
    if (!view) return [];
    const wanted = String(number);
    return Array.from(view.querySelectorAll('.leaflet-marker-pane > *, .leaflet-tooltip-pane > *, .leaflet-marker-icon, [data-plant-index], [data-impianto-index], [data-marker-number]')).filter((node) => {
      const vals = [node.dataset?.plantIndex,node.dataset?.impiantoIndex,node.dataset?.markerNumber].filter((v) => v != null).map(String);
      return vals.includes(wanted)
        || vals.includes(String(Number(wanted) - 1))
        || String(node.textContent || '').trim() === wanted;
    });
  }

  function flashMatches(matches) {
    clearFlashes();
    clearSelected();
    matches.slice(0, MAX_RESULTS).forEach(({item,index}) => {
      markerCandidatesForNumber(plantNumber(item,index)).forEach((node) => node.classList.add(FLASH_CLASS));
    });
  }

  function centerMatches(matches) {
    if (!matches.length || matches.length > MAX_AUTO_CENTER) return;
    const coords = matches.map(({item}) => plantCoordinates(item)).filter(Boolean);
    const mapInstance = getFullscreenMap();
    if (!coords.length || !mapInstance) return;
    try {
      if (coords.length === 1) {
        mapInstance.setView(coords[0], Math.max(Number(mapInstance.getZoom?.() || 0), 16), { animate: true });
      } else if (window.L?.latLngBounds) {
        mapInstance.fitBounds(window.L.latLngBounds(coords), { padding: [60,60], maxZoom: 16, animate: true });
      } else {
        mapInstance.fitBounds(coords, { padding: [60,60], maxZoom: 16, animate: true });
      }
    } catch (error) {
      console.warn('Centratura risultati ricerca non completata:', error);
    }
  }

  function ensureSuggestions(input) {
    let box = document.getElementById(SUGGESTIONS_ID);
    if (box) return box;
    const parent = input.closest('.hera-map-search-shell') || input.parentElement;
    if (!parent) return null;
    if (getComputedStyle(parent).position === 'static') parent.style.position = 'relative';
    box = document.createElement('div');
    box.id = SUGGESTIONS_ID;
    box.className = 'hera-map-live-suggestions';
    box.hidden = true;
    box.setAttribute('role','listbox');
    parent.appendChild(box);
    return box;
  }

  function hideSuggestions() {
    const box = document.getElementById(SUGGESTIONS_ID);
    if (box) { box.hidden = true; box.textContent = ''; }
  }

  function selectMatch(match, input) {
    clearFlashes();
    clearSelected();
    hideSuggestions();
    input.value = plantName(match.item, match.index);
    const nodes = markerCandidatesForNumber(plantNumber(match.item, match.index));
    nodes.forEach((node) => node.classList.add(SELECTED_CLASS));
    centerMatches([match]);
    const clickable = nodes.find((node) => typeof node.click === 'function');
    try { clickable?.click(); } catch (_) {}
  }

  function renderSuggestions(matches, input) {
    const box = ensureSuggestions(input);
    if (!box) return;
    box.textContent = '';
    const visible = matches.slice(0, MAX_RESULTS);
    if (!visible.length) {
      const empty = document.createElement('div');
      empty.className = 'hera-map-live-empty';
      empty.textContent = 'Nessun impianto trovato';
      box.appendChild(empty);
      box.hidden = false;
      return;
    }
    visible.forEach((match) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'hera-map-live-suggestion';
      button.dataset.plantIndex = String(match.index);
      const title = document.createElement('strong');
      title.textContent = `${plantNumber(match.item, match.index)} · ${plantName(match.item, match.index)}`;
      const meta = document.createElement('span');
      meta.textContent = [plantComune(match.item), plantAddress(match.item), plantCode(match.item) ? `ID SAP ${plantCode(match.item)}` : ''].filter(Boolean).join(' • ');
      button.append(title, meta);
      button.addEventListener('click', () => selectMatch(match, input));
      box.appendChild(button);
    });
    box.hidden = false;
  }

  let searchTimer = null;
  function handleSearchInput(input) {
    window.clearTimeout(searchTimer);
    searchTimer = window.setTimeout(() => {
      const query = normalize(input.value);
      if (!query) {
        clearFlashes();
        clearSelected();
        hideSuggestions();
        return;
      }
      const matches = findMatches(query);
      renderSuggestions(matches, input);
      if (!matches.length) {
        clearFlashes();
        clearSelected();
        return;
      }
      const bestScore = matches[0].score;
      const best = matches.filter((entry) => entry.score === bestScore).slice(0, MAX_RESULTS);
      flashMatches(best);
      centerMatches(best);
    }, 100);
  }

  function hideDrawUiSafely() {
    const page = document.getElementById(PAGE_ID);
    if (!page) return;
    ['#map-fullscreen-draw-btn','#map-draw-btn','#draw-map-btn','#map-fullscreen-undo-btn','#map-fullscreen-redo-btn','#map-fullscreen-clear-btn','.map-draw-toolbar','.map-drawing-toolbar','.leaflet-draw','.leaflet-draw-toolbar','[data-map-draw]','[data-draw-map]']
      .forEach((selector) => page.querySelectorAll(selector).forEach((node) => node.remove()));
  }

  function installStyles() {
    if (document.getElementById('hera-map-search-focus-style-v3')) return;
    const style = document.createElement('style');
    style.id = 'hera-map-search-focus-style-v3';
    style.textContent = `
      @keyframes heraSearchPulseV3 {
        0%,100% { transform:scale(1); filter:brightness(1.08) drop-shadow(0 0 5px rgba(239,68,68,.98)) drop-shadow(0 0 12px rgba(239,68,68,.90)); }
        50% { transform:scale(1.32); filter:brightness(1.24) drop-shadow(0 0 6px rgba(37,99,235,1)) drop-shadow(0 0 14px rgba(37,99,235,.98)); }
      }
      #${PAGE_ID} .${FLASH_CLASS} { animation:heraSearchPulseV3 .72s ease-in-out infinite !important; transform-origin:center center !important; z-index:9999 !important; }
      #${PAGE_ID} .${SELECTED_CLASS} { filter:brightness(1.12) drop-shadow(0 0 6px rgba(37,99,235,1)) drop-shadow(0 0 14px rgba(37,99,235,.8)) !important; transform:scale(1.14) !important; transform-origin:center center !important; z-index:9998 !important; }
      #${PAGE_ID} .hera-map-live-suggestions { position:absolute; left:0; right:0; top:calc(100% + 8px); z-index:5000; max-height:min(52vh,420px); overflow:auto; padding:7px; border:1px solid rgba(15,23,42,.12); border-radius:18px; background:rgba(255,255,255,.995); box-shadow:0 18px 40px rgba(15,23,42,.24); -webkit-overflow-scrolling:touch; }
      #${PAGE_ID} .hera-map-live-suggestions[hidden] { display:none !important; }
      #${PAGE_ID} .hera-map-live-suggestion { display:flex; flex-direction:column; gap:3px; width:100%; padding:11px 12px; border:0; border-radius:13px; background:transparent; text-align:left; color:#111827; }
      #${PAGE_ID} .hera-map-live-suggestion:active { background:#eef4ff; }
      #${PAGE_ID} .hera-map-live-suggestion strong { font-size:14px; font-weight:800; line-height:1.25; }
      #${PAGE_ID} .hera-map-live-suggestion span { font-size:12px; font-weight:600; color:#64748b; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
      #${PAGE_ID} .hera-map-live-empty { padding:14px; font-size:14px; font-weight:700; color:#64748b; }
      #${PAGE_ID} .leaflet-draw,#${PAGE_ID} .leaflet-draw-toolbar,#${PAGE_ID} [data-map-draw],#${PAGE_ID} [data-draw-map] { display:none !important; }
    `;
    document.head.appendChild(style);
  }

  function setupInput() {
    const input = document.getElementById(INPUT_ID);
    if (!input) return;
    input.placeholder = 'Cerca impianto, numero, comune, indirizzo o ID SAP';
    input.setAttribute('aria-label','Cerca impianto per nome, numero, comune, indirizzo o ID SAP');
    ensureSuggestions(input);
    if (input.dataset.searchFocusEnhancementV3 === '1') return;
    input.dataset.searchFocusEnhancementV3 = '1';
    input.addEventListener('input', () => handleSearchInput(input));
    input.addEventListener('search', () => handleSearchInput(input));
    input.addEventListener('focus', () => { if (normalize(input.value)) handleSearchInput(input); });
  }

  function setupSelectionStop() {
    const page = document.getElementById(PAGE_ID);
    if (!page || page.dataset.searchSelectionStopV3 === '1') return;
    page.dataset.searchSelectionStopV3 = '1';
    page.addEventListener('click', (event) => {
      if (event.target.closest?.(`#${SUGGESTIONS_ID}`)) return;
      if (event.target.closest?.('.hera-map-search-clear')) {
        clearFlashes(); clearSelected(); hideSuggestions(); return;
      }
      const marker = event.target.closest?.(`#${VIEW_ID} .leaflet-marker-icon, #${VIEW_ID} .leaflet-tooltip-pane > *, #${VIEW_ID} [data-plant-index], #${VIEW_ID} [data-impianto-index], #${VIEW_ID} [data-marker-number]`);
      if (marker) {
        clearFlashes(); clearSelected(); hideSuggestions(); marker.classList.add(SELECTED_CLASS);
      }
    }, true);
  }

  function ensureMapVisible() {
    const view = document.getElementById(VIEW_ID);
    if (!view) return;
    view.style.removeProperty('display');
    view.style.removeProperty('visibility');
    view.style.removeProperty('opacity');
    try { getFullscreenMap()?.invalidateSize?.({ animate:false }); } catch (_) {}
  }

  function apply() {
    installStyles();
    hideDrawUiSafely();
    setupInput();
    setupSelectionStop();
    ensureMapVisible();
  }

  const observer = new MutationObserver(() => window.requestAnimationFrame(apply));
  observer.observe(document.documentElement, { childList:true, subtree:true });
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', apply, { once:true });
  else apply();
})();

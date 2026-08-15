(() => {
  'use strict';

  if (window.__heraMapSearchFocusInstalled) return;
  window.__heraMapSearchFocusInstalled = true;

  const INPUT_ID = 'map-fullscreen-number-search-input';
  const PAGE_ID = 'map-fullscreen-page';
  const MAX_AUTO_CENTER_MATCHES = 6;
  const MAX_FLASH_MATCHES = 8;
  const FLASH_CLASS = 'hera-search-match-flash';

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

  function plantCoordinates(item) {
    const latRaw = firstValue(item, [
      'lat', 'latitude', 'Latitudine', 'latitudine', 'GPS(Y)', 'gpsY', 'gps_y',
      'coordinateY', 'coordinate_y', 'y'
    ]);
    const lngRaw = firstValue(item, [
      'lng', 'lon', 'longitude', 'Longitudine', 'longitudine', 'GPS(X)', 'gpsX', 'gps_x',
      'coordinateX', 'coordinate_x', 'x'
    ]);

    const lat = Number(String(latRaw).replace(',', '.'));
    const lng = Number(String(lngRaw).replace(',', '.'));
    if (Number.isFinite(lat) && Number.isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180) {
      return [lat, lng];
    }

    const combined = firstValue(item, [
      'coordinate', 'Coordinate', 'coordinateGPS', 'Coordinate GPS', 'gps', 'GPS'
    ]);
    if (combined) {
      const nums = String(combined).match(/-?\d+(?:[.,]\d+)?/g)?.map((value) => Number(value.replace(',', '.'))) || [];
      if (nums.length >= 2) {
        const a = nums[0];
        const b = nums[1];
        if (Math.abs(a) <= 90 && Math.abs(b) <= 180) return [a, b];
        if (Math.abs(b) <= 90 && Math.abs(a) <= 180) return [b, a];
      }
    }
    return null;
  }

  function tokensFor(item, index) {
    return {
      name: normalize(plantName(item, index)),
      comune: normalize(plantComune(item)),
      address: normalize(plantAddress(item)),
      code: normalize(plantCode(item)),
      number: normalize(String(plantNumber(item, index)))
    };
  }

  function matchScore(item, index, query) {
    const tokens = tokensFor(item, index);
    const values = Object.values(tokens).filter(Boolean);
    if (!values.length) return Infinity;
    if (tokens.number === query || tokens.code === query || tokens.name === query) return 0;
    if (values.some((value) => value.startsWith(query))) return 1;
    if (values.some((value) => value.split(/\s+/).some((word) => word.startsWith(query)))) return 2;
    if (values.some((value) => value.includes(query))) return 3;
    return Infinity;
  }

  function findMatches(query) {
    return getCurrentImpianti()
      .map((item, index) => ({ item, index, score: matchScore(item, index, query) }))
      .filter((entry) => Number.isFinite(entry.score))
      .sort((a, b) => a.score - b.score || plantNumber(a.item, a.index) - plantNumber(b.item, b.index));
  }

  function getFullscreenMap() {
    const candidates = [
      window.mapFullscreen, window.fullscreenMap, window.mapFullScreen,
      window.fullScreenMap, window.mapFull, window.map
    ].filter(Boolean);
    for (const candidate of candidates) {
      try {
        if (typeof candidate?.fitBounds !== 'function' || typeof candidate?.setView !== 'function') continue;
        const container = candidate.getContainer?.();
        if (!container) continue;
        if (container.id === 'map-fullscreen-view' || container.closest?.(`#${PAGE_ID}`)) return candidate;
      } catch (_) {}
    }

    try {
      for (const key of Object.keys(window)) {
        let candidate;
        try { candidate = window[key]; } catch (_) { continue; }
        if (!candidate || typeof candidate !== 'object') continue;
        if (typeof candidate.fitBounds !== 'function' || typeof candidate.setView !== 'function') continue;
        const container = candidate.getContainer?.();
        if (container?.id === 'map-fullscreen-view' || container?.closest?.(`#${PAGE_ID}`)) return candidate;
      }
    } catch (_) {}
    return null;
  }

  function clearFlashes() {
    document.querySelectorAll(`.${FLASH_CLASS}`).forEach((node) => node.classList.remove(FLASH_CLASS));
  }

  function markerCandidatesForNumber(number) {
    const wanted = String(number);
    const page = document.getElementById(PAGE_ID) || document;
    const markerNodes = Array.from(page.querySelectorAll(
      '#map-fullscreen-view .leaflet-marker-pane > *, #map-fullscreen-view .leaflet-tooltip-pane > *, #map-fullscreen-view .leaflet-marker-icon, #map-fullscreen-view [data-plant-index], #map-fullscreen-view [data-impianto-index], #map-fullscreen-view [data-marker-number]'
    ));
    return markerNodes.filter((node) => {
      const datasetValues = [node.dataset?.plantIndex, node.dataset?.impiantoIndex, node.dataset?.markerNumber]
        .filter((value) => value != null)
        .map(String);
      if (datasetValues.includes(wanted) || datasetValues.includes(String(Number(wanted) - 1))) return true;
      return String(node.textContent || '').trim() === wanted;
    });
  }

  function flashMatches(matches) {
    clearFlashes();
    matches.slice(0, MAX_FLASH_MATCHES).forEach(({ item, index }) => {
      const number = plantNumber(item, index);
      markerCandidatesForNumber(number).forEach((node) => node.classList.add(FLASH_CLASS));
    });
  }

  function centerMatches(matches) {
    if (!matches.length || matches.length > MAX_AUTO_CENTER_MATCHES) return;
    const coords = matches.map(({ item }) => plantCoordinates(item)).filter(Boolean);
    if (!coords.length) return;
    const mapInstance = getFullscreenMap();
    if (!mapInstance) return;

    try {
      if (coords.length === 1) {
        const currentZoom = Number(mapInstance.getZoom?.() || 0);
        mapInstance.setView(coords[0], Math.max(currentZoom, 16), { animate: true });
        return;
      }
      if (window.L?.latLngBounds) {
        mapInstance.fitBounds(window.L.latLngBounds(coords), {
          padding: [54, 54],
          maxZoom: 16,
          animate: true
        });
      } else {
        mapInstance.fitBounds(coords, { padding: [54, 54], maxZoom: 16, animate: true });
      }
    } catch (error) {
      console.warn('Centratura risultati ricerca non completata:', error);
    }
  }

  let searchTimer = null;
  function handleSearchInput(input) {
    window.clearTimeout(searchTimer);
    searchTimer = window.setTimeout(() => {
      const query = normalize(input.value);
      if (!query) {
        clearFlashes();
        return;
      }
      const matches = findMatches(query);
      if (!matches.length) {
        clearFlashes();
        return;
      }
      const bestScore = matches[0].score;
      const nearest = matches.filter((entry) => entry.score === bestScore).slice(0, MAX_FLASH_MATCHES);
      flashMatches(nearest);
      centerMatches(nearest);
    }, 160);
  }

  function hideDrawUi() {
    const page = document.getElementById(PAGE_ID);
    if (!page) return;

    [
      '#map-fullscreen-draw-btn', '#map-draw-btn', '#draw-map-btn',
      '#map-fullscreen-undo-btn', '#map-fullscreen-redo-btn', '#map-fullscreen-clear-btn',
      '.map-draw-toolbar', '.map-drawing-toolbar', '.leaflet-draw', '.leaflet-draw-toolbar',
      '[data-map-draw]', '[data-draw-map]'
    ].forEach((selector) => page.querySelectorAll(selector).forEach((node) => node.remove()));

    page.querySelectorAll('button, [role="button"], .leaflet-control, .map-fullscreen-hint, .map-hint, p, div').forEach((node) => {
      if (!node.isConnected) return;
      const signature = normalize([
        node.textContent,
        node.getAttribute?.('title'),
        node.getAttribute?.('aria-label'),
        node.id,
        node.className
      ].filter(Boolean).join(' '));
      const isDrawControl = /(^|\s)(disegna|disegno|drawing|draw)(\s|$)/.test(signature)
        || /usa.*disegna.*per.*perimetro/.test(signature)
        || /annulla.*disegno|ripristina.*disegno|cancella.*disegno|condividi.*disegno/.test(signature);
      if (!isDrawControl) return;
      if (node.closest?.('.hera-map-search-results')) return;
      if (node.matches?.('div') && node.children.length > 5) return;
      node.remove();
    });
  }

  function installStyles() {
    if (document.getElementById('hera-map-search-focus-style')) return;
    const style = document.createElement('style');
    style.id = 'hera-map-search-focus-style';
    style.textContent = `
      @keyframes heraSearchPulse {
        0%, 100% { transform: scale(1); filter: brightness(1); }
        50% { transform: scale(1.28); filter: brightness(1.22) drop-shadow(0 0 9px rgba(37, 99, 235, .75)); }
      }
      #${PAGE_ID} .${FLASH_CLASS} {
        animation: heraSearchPulse .72s ease-in-out infinite !important;
        transform-origin: center center !important;
        z-index: 9999 !important;
      }
      #${PAGE_ID} .leaflet-draw,
      #${PAGE_ID} .leaflet-draw-toolbar,
      #${PAGE_ID} [data-map-draw],
      #${PAGE_ID} [data-draw-map] { display: none !important; }
    `;
    document.head.appendChild(style);
  }

  function setupInput() {
    const input = document.getElementById(INPUT_ID);
    if (!input) return;
    input.placeholder = 'Cerca impianto, numero, comune, indirizzo o ID SAP';
    input.setAttribute('aria-label', 'Cerca impianto per nome, numero, comune, indirizzo o ID SAP');
    if (input.dataset.searchFocusEnhancement === '1') return;
    input.dataset.searchFocusEnhancement = '1';
    input.addEventListener('input', () => handleSearchInput(input));
    input.addEventListener('search', () => handleSearchInput(input));
  }

  function apply() {
    installStyles();
    hideDrawUi();
    setupInput();
  }

  const observer = new MutationObserver(apply);
  observer.observe(document.documentElement, { childList: true, subtree: true });

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', apply, { once: true });
  else apply();
})();

(() => {
  'use strict';

  if (window.__heraMapSearchFocusV2Installed) return;
  window.__heraMapSearchFocusV2Installed = true;

  const INPUT_ID = 'map-fullscreen-number-search-input';
  const PAGE_ID = 'map-fullscreen-page';
  const VIEW_ID = 'map-fullscreen-view';
  const MAX_AUTO_CENTER_MATCHES = 6;
  const MAX_FLASH_MATCHES = 8;
  const FLASH_CLASS = 'hera-search-match-flash';

  const normalize = (value) => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
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
    const values = [
      normalize(plantName(item,index)), normalize(plantComune(item)), normalize(plantAddress(item)),
      normalize(plantCode(item)), normalize(String(plantNumber(item,index)))
    ].filter(Boolean);
    if (!values.length) return Infinity;
    if (values.some((v) => v === query)) return 0;
    if (values.some((v) => v.startsWith(query))) return 1;
    if (values.some((v) => v.split(/\s+/).some((w) => w.startsWith(query)))) return 2;
    if (values.some((v) => v.includes(query))) return 3;
    return Infinity;
  }

  function findMatches(query) {
    return getCurrentImpianti().map((item,index) => ({item,index,score:matchScore(item,index,query)}))
      .filter((e) => Number.isFinite(e.score))
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

  function clearFlashes() { document.querySelectorAll(`.${FLASH_CLASS}`).forEach((n) => n.classList.remove(FLASH_CLASS)); }
  function markerCandidatesForNumber(number) {
    const view = document.getElementById(VIEW_ID);
    if (!view) return [];
    const wanted = String(number);
    return Array.from(view.querySelectorAll('.leaflet-marker-pane > *, .leaflet-tooltip-pane > *, .leaflet-marker-icon, [data-plant-index], [data-impianto-index], [data-marker-number]')).filter((node) => {
      const vals = [node.dataset?.plantIndex,node.dataset?.impiantoIndex,node.dataset?.markerNumber].filter((v) => v != null).map(String);
      return vals.includes(wanted) || vals.includes(String(Number(wanted)-1)) || String(node.textContent || '').trim() === wanted;
    });
  }
  function flashMatches(matches) {
    clearFlashes();
    matches.slice(0,MAX_FLASH_MATCHES).forEach(({item,index}) => markerCandidatesForNumber(plantNumber(item,index)).forEach((n) => n.classList.add(FLASH_CLASS)));
  }
  function centerMatches(matches) {
    if (!matches.length || matches.length > MAX_AUTO_CENTER_MATCHES) return;
    const coords = matches.map(({item}) => plantCoordinates(item)).filter(Boolean);
    const mapInstance = getFullscreenMap();
    if (!coords.length || !mapInstance) return;
    try {
      if (coords.length === 1) mapInstance.setView(coords[0], Math.max(Number(mapInstance.getZoom?.() || 0),16), {animate:true});
      else if (window.L?.latLngBounds) mapInstance.fitBounds(window.L.latLngBounds(coords), {padding:[54,54],maxZoom:16,animate:true});
      else mapInstance.fitBounds(coords, {padding:[54,54],maxZoom:16,animate:true});
    } catch (error) { console.warn('Centratura risultati ricerca non completata:', error); }
  }

  let searchTimer = null;
  function handleSearchInput(input) {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      const query = normalize(input.value);
      if (!query) { clearFlashes(); return; }
      const matches = findMatches(query);
      if (!matches.length) { clearFlashes(); return; }
      const bestScore = matches[0].score;
      const best = matches.filter((e) => e.score === bestScore).slice(0,MAX_FLASH_MATCHES);
      flashMatches(best);
      centerMatches(best);
    },160);
  }

  function hideDrawUiSafely() {
    const page = document.getElementById(PAGE_ID);
    if (!page) return;
    ['#map-fullscreen-draw-btn','#map-draw-btn','#draw-map-btn','#map-fullscreen-undo-btn','#map-fullscreen-redo-btn','#map-fullscreen-clear-btn','.map-draw-toolbar','.map-drawing-toolbar','.leaflet-draw','.leaflet-draw-toolbar','[data-map-draw]','[data-draw-map]']
      .forEach((selector) => page.querySelectorAll(selector).forEach((node) => node.remove()));

    // IMPORTANTE: non includere DIV generici. Il vecchio runtime poteva rimuovere
    // il contenitore della mappa perché il suo textContent conteneva il messaggio Disegna.
    page.querySelectorAll('button, [role="button"], .map-fullscreen-hint, .map-hint, p').forEach((node) => {
      if (!node.isConnected) return;
      const view = document.getElementById(VIEW_ID);
      const input = document.getElementById(INPUT_ID);
      if (node === view || (view && node.contains(view)) || (input && node.contains(input))) return;
      const signature = normalize([node.textContent,node.getAttribute?.('title'),node.getAttribute?.('aria-label'),node.id,node.className].filter(Boolean).join(' '));
      if (/(^|\s)(disegna|disegno|drawing|draw)(\s|$)/.test(signature) || /usa.*disegna.*per.*perimetro/.test(signature) || /annulla.*disegno|ripristina.*disegno|cancella.*disegno|condividi.*disegno/.test(signature)) node.remove();
    });
  }

  function installStyles() {
    if (document.getElementById('hera-map-search-focus-style-v2')) return;
    const style = document.createElement('style');
    style.id = 'hera-map-search-focus-style-v2';
    style.textContent = `
      @keyframes heraSearchPulseV2 { 0%,100%{transform:scale(1);filter:brightness(1)} 50%{transform:scale(1.28);filter:brightness(1.22) drop-shadow(0 0 9px rgba(37,99,235,.75))} }
      #${PAGE_ID} .${FLASH_CLASS}{animation:heraSearchPulseV2 .72s ease-in-out infinite !important;transform-origin:center center !important;z-index:9999 !important}
      #${PAGE_ID} .leaflet-draw,#${PAGE_ID} .leaflet-draw-toolbar,#${PAGE_ID} [data-map-draw],#${PAGE_ID} [data-draw-map]{display:none !important}
    `;
    document.head.appendChild(style);
  }

  function setupInput() {
    const input = document.getElementById(INPUT_ID);
    if (!input) return;
    input.placeholder = 'Cerca impianto, numero, comune, indirizzo o ID SAP';
    input.setAttribute('aria-label','Cerca impianto per nome, numero, comune, indirizzo o ID SAP');
    if (input.dataset.searchFocusEnhancementV2 === '1') return;
    input.dataset.searchFocusEnhancementV2 = '1';
    input.addEventListener('input', () => handleSearchInput(input));
    input.addEventListener('search', () => handleSearchInput(input));
  }

  function ensureMapVisible() {
    const view = document.getElementById(VIEW_ID);
    if (!view) return;
    view.style.removeProperty('display');
    view.style.removeProperty('visibility');
    view.style.removeProperty('opacity');
    try { getFullscreenMap()?.invalidateSize?.({animate:false}); } catch (_) {}
  }

  function apply() { installStyles(); hideDrawUiSafely(); setupInput(); ensureMapVisible(); }
  const observer = new MutationObserver(() => requestAnimationFrame(apply));
  observer.observe(document.documentElement,{childList:true,subtree:true});
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded',apply,{once:true}); else apply();
})();

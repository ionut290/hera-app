(() => {
  'use strict';

  if (window.__heraSafeFullscreenSearchV4Installed) return;
  window.__heraSafeFullscreenSearchV4Installed = true;

  const INPUT_ID = 'map-fullscreen-number-search-input';
  const PAGE_ID = 'map-fullscreen-page';
  const VIEW_ID = 'map-fullscreen-view';
  const BOX_ID = 'hera-safe-map-suggestions';
  const MAX_RESULTS = 8;
  const MAX_AUTO_CENTER = 6;

  let overlayGroup = null;
  let blinkTimer = null;
  let blinkBlue = false;
  let boundInput = null;

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

  function plantComune(item) {
    return firstValue(item, ['comune','Comune','ubicazione','localita','località','city']);
  }

  function plantAddress(item) {
    return firstValue(item, ['descrizioneVia','descrizione_via','Descrizione via','indirizzo','Indirizzo','via','address']);
  }

  function plantCode(item) {
    return firstValue(item, ['idSap','idSAP','ID SAP','idsap','sap','codiceSap','codiceSAP','codice','code','id']);
  }

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
      normalize(plantName(item,index)),
      normalize(plantComune(item)),
      normalize(plantAddress(item)),
      normalize(plantCode(item)),
      normalize(String(plantNumber(item,index)))
    ].filter(Boolean);

    if (!values.length) return Infinity;
    if (values.some((value) => value === query)) return 0;
    if (values.some((value) => value.startsWith(query))) return 1;
    if (values.some((value) => value.split(/\s+/).some((word) => word.startsWith(query)))) return 2;
    if (values.some((value) => value.includes(query))) return 3;
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

  function ensureStyles() {
    if (document.getElementById('hera-safe-map-search-v4-style')) return;
    const style = document.createElement('style');
    style.id = 'hera-safe-map-search-v4-style';
    style.textContent = `
      #${PAGE_ID} .leaflet-draw,
      #${PAGE_ID} .leaflet-draw-toolbar,
      #${PAGE_ID} [data-map-draw],
      #${PAGE_ID} [data-draw-map],
      #${PAGE_ID} #map-fullscreen-draw-btn,
      #${PAGE_ID} #map-draw-btn,
      #${PAGE_ID} #draw-map-btn,
      #${PAGE_ID} .map-fullscreen-hint,
      #${PAGE_ID} .map-hint { display:none !important; }
      #${BOX_ID} {
        position:fixed;
        z-index:2147483000;
        max-height:min(52vh,420px);
        overflow:auto;
        padding:7px;
        border:1px solid rgba(15,23,42,.12);
        border-radius:16px;
        background:rgba(255,255,255,.995);
        box-shadow:0 16px 38px rgba(15,23,42,.24);
        -webkit-overflow-scrolling:touch;
      }
      #${BOX_ID}[hidden] { display:none !important; }
      #${BOX_ID} button {
        display:flex;
        flex-direction:column;
        gap:3px;
        width:100%;
        padding:11px 12px;
        border:0;
        border-radius:12px;
        background:transparent;
        text-align:left;
        color:#111827;
      }
      #${BOX_ID} button:active { background:#eef4ff; }
      #${BOX_ID} strong { font-size:14px; font-weight:800; }
      #${BOX_ID} span { font-size:12px; font-weight:600; color:#64748b; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
      #${BOX_ID} .hera-empty { padding:14px; color:#64748b; font-size:14px; font-weight:700; }
    `;
    document.head.appendChild(style);
  }

  function ensureBox(input) {
    let box = document.getElementById(BOX_ID);
    if (!box) {
      box = document.createElement('div');
      box.id = BOX_ID;
      box.hidden = true;
      document.body.appendChild(box);
    }
    positionBox(input, box);
    return box;
  }

  function positionBox(input, box) {
    if (!input || !box) return;
    const rect = input.getBoundingClientRect();
    box.style.left = `${Math.max(8, rect.left)}px`;
    box.style.top = `${Math.min(window.innerHeight - 80, rect.bottom + 8)}px`;
    box.style.width = `${Math.max(220, Math.min(rect.width, window.innerWidth - 16))}px`;
  }

  function hideBox() {
    const box = document.getElementById(BOX_ID);
    if (box) {
      box.hidden = true;
      box.textContent = '';
    }
  }

  function clearSearchOverlay() {
    if (blinkTimer) {
      clearInterval(blinkTimer);
      blinkTimer = null;
    }
    const map = getFullscreenMap();
    try {
      if (overlayGroup && map?.hasLayer?.(overlayGroup)) map.removeLayer(overlayGroup);
    } catch (_) {}
    overlayGroup = null;
    blinkBlue = false;
  }

  function drawSearchOverlay(matches) {
    clearSearchOverlay();
    const map = getFullscreenMap();
    if (!map || !window.L?.layerGroup || !window.L?.circleMarker) return;

    const coords = matches.map((match) => ({ match, coords: plantCoordinates(match.item) })).filter((entry) => entry.coords);
    if (!coords.length) return;

    overlayGroup = window.L.layerGroup().addTo(map);
    const circles = coords.map(({coords}) => window.L.circleMarker(coords, {
      radius: 13,
      weight: 4,
      color: '#ef4444',
      fillColor: '#ef4444',
      fillOpacity: 0.22,
      opacity: 0.95,
      interactive: false,
      pane: 'markerPane'
    }).addTo(overlayGroup));

    blinkTimer = setInterval(() => {
      blinkBlue = !blinkBlue;
      const color = blinkBlue ? '#2563eb' : '#ef4444';
      circles.forEach((circle) => circle.setStyle({ color, fillColor: color, radius: blinkBlue ? 16 : 13 }));
    }, 430);
  }

  function centerMatches(matches) {
    if (!matches.length || matches.length > MAX_AUTO_CENTER) return;
    const coords = matches.map((match) => plantCoordinates(match.item)).filter(Boolean);
    const map = getFullscreenMap();
    if (!map || !coords.length) return;
    try {
      if (coords.length === 1) map.setView(coords[0], Math.max(Number(map.getZoom?.() || 0), 16), { animate:true });
      else if (window.L?.latLngBounds) map.fitBounds(window.L.latLngBounds(coords), { padding:[60,60], maxZoom:16, animate:true });
    } catch (_) {}
  }

  function openPlantCard(match) {
    const map = getFullscreenMap();
    const coords = plantCoordinates(match?.item);
    if (!map || !coords) return;

    const targetLat = Number(coords[0]);
    const targetLng = Number(coords[1]);
    let bestPopupLayer = null;
    let bestPopupDistance = Infinity;
    let bestClickableLayer = null;
    let bestClickableDistance = Infinity;

    try {
      map.eachLayer((layer) => {
        if (!layer || typeof layer.getLatLng !== 'function') return;
        let latLng;
        try { latLng = layer.getLatLng(); } catch (_) { return; }
        const lat = Number(latLng?.lat);
        const lng = Number(latLng?.lng);
        if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;

        const distance = ((lat - targetLat) ** 2) + ((lng - targetLng) ** 2);
        const interactive = layer.options?.interactive !== false;
        if (!interactive) return;

        const popup = typeof layer.getPopup === 'function' ? layer.getPopup() : null;
        if (popup && typeof layer.openPopup === 'function' && distance < bestPopupDistance) {
          bestPopupLayer = layer;
          bestPopupDistance = distance;
        }

        if (typeof layer.fire === 'function' && distance < bestClickableDistance) {
          bestClickableLayer = layer;
          bestClickableDistance = distance;
        }
      });
    } catch (_) {}

    const zoom = Math.max(Number(map.getZoom?.() || 0), 17);
    try { map.setView(coords, zoom, { animate:true }); } catch (_) {}

    window.setTimeout(() => {
      try {
        if (bestPopupLayer) {
          bestPopupLayer.openPopup();
          return;
        }
        bestClickableLayer?.fire?.('click', { latlng: bestClickableLayer.getLatLng?.() });
      } catch (_) {}
    }, 260);
  }

  function renderSuggestions(matches, input) {
    const box = ensureBox(input);
    box.textContent = '';
    const visible = matches.slice(0, MAX_RESULTS);
    if (!visible.length) {
      const empty = document.createElement('div');
      empty.className = 'hera-empty';
      empty.textContent = 'Nessun impianto trovato';
      box.appendChild(empty);
      box.hidden = false;
      return;
    }

    visible.forEach((match) => {
      const button = document.createElement('button');
      button.type = 'button';
      const title = document.createElement('strong');
      title.textContent = `${plantNumber(match.item,match.index)} · ${plantName(match.item,match.index)}`;
      const meta = document.createElement('span');
      meta.textContent = [plantComune(match.item), plantAddress(match.item), plantCode(match.item) ? `ID SAP ${plantCode(match.item)}` : ''].filter(Boolean).join(' • ');
      button.append(title, meta);
      button.addEventListener('click', () => {
        input.value = plantName(match.item,match.index);
        hideBox();
        clearSearchOverlay();
        openPlantCard(match);
      });
      box.appendChild(button);
    });

    box.hidden = false;
  }

  let searchDebounce = null;
  function onSearch(input) {
    clearTimeout(searchDebounce);
    searchDebounce = setTimeout(() => {
      const query = normalize(input.value);
      if (!query) {
        hideBox();
        clearSearchOverlay();
        return;
      }
      const matches = findMatches(query);
      renderSuggestions(matches, input);
      if (!matches.length) {
        clearSearchOverlay();
        return;
      }
      const bestScore = matches[0].score;
      const best = matches.filter((entry) => entry.score === bestScore).slice(0, MAX_RESULTS);
      drawSearchOverlay(best);
      centerMatches(best);
    }, 110);
  }

  function bindInput(input) {
    if (!input || input === boundInput) return;
    boundInput = input;
    input.type = 'search';
    input.inputMode = 'search';
    input.placeholder = 'Cerca impianto, numero, comune, indirizzo o ID SAP';
    input.setAttribute('autocomplete','off');
    input.addEventListener('input', () => onSearch(input));
    input.addEventListener('search', () => onSearch(input));
    input.addEventListener('focus', () => { if (normalize(input.value)) onSearch(input); });
    input.addEventListener('blur', () => setTimeout(() => {
      const box = document.getElementById(BOX_ID);
      if (box && !box.matches(':hover')) hideBox();
    }, 180));
  }

  function initWhenAvailable() {
    ensureStyles();
    const input = document.getElementById(INPUT_ID);
    if (input) bindInput(input);
  }

  window.addEventListener('resize', () => {
    const box = document.getElementById(BOX_ID);
    if (boundInput && box && !box.hidden) positionBox(boundInput, box);
  }, { passive:true });

  window.addEventListener('scroll', () => {
    const box = document.getElementById(BOX_ID);
    if (boundInput && box && !box.hidden) positionBox(boundInput, box);
  }, { passive:true });

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initWhenAvailable, { once:true });
  } else {
    initWhenAvailable();
  }

  let tries = 0;
  const initTimer = setInterval(() => {
    tries += 1;
    initWhenAvailable();
    if (boundInput || tries >= 40) clearInterval(initTimer);
  }, 500);
})();

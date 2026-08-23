(() => {
  'use strict';
  if (window.HeraStreetViewCards?.installed) return;

  const VERSION = '1.1.0';
  const METADATA_CACHE_KEY = 'heraStreetViewMetadataV1';
  const METADATA_TTL = 7 * 24 * 60 * 60 * 1000;
  const text = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();
  const upper = (value) => text(value).toLocaleUpperCase('it-IT');

  function resolveApiKey() {
    const globals = [window.HERA_GOOGLE_MAPS_API_KEY, window.GOOGLE_MAPS_API_KEY, window.GOOGLE_API_KEY, window.mapsApiKey, window.googleMapsApiKey, window.__GOOGLE_MAPS_API_KEY__]
      .map(text).filter(Boolean);
    if (globals.length) return globals[0];
    try { return text(window.firebase?.app?.()?.options?.apiKey); } catch (_) {}
    try { return text(window.firebaseApp?.options?.apiKey); } catch (_) {}
    return '';
  }

  function getPlants() {
    try { if (Array.isArray(currentImpianti)) return currentImpianti; } catch (_) {}
    return Array.isArray(window.currentImpianti) ? window.currentImpianti : [];
  }

  function getIds(item) {
    try {
      if (typeof getImpiantoDocIds === 'function') return getImpiantoDocIds(item).map(text).filter(Boolean);
    } catch (_) {}
    return [item?.physicalPlantId, item?.impiantoId, item?.id, item?.idSap, item?.idSAP, item?.sap].map(text).filter(Boolean);
  }

  function findPlant(nav, card) {
    const key = text(nav.dataset?.impiantoKey || nav.dataset?.impiantoId || card.dataset?.impiantoKey || card.dataset?.impiantoId);
    const name = text(nav.dataset?.impiantoName);
    const plants = getPlants();
    if (key) {
      const byId = plants.find((item) => getIds(item).includes(key));
      if (byId) return byId;
    }
    if (name) {
      const target = upper(name);
      const byName = plants.find((item) => [item?.denominazione, item?.nome, item?.impianto].map(upper).includes(target));
      if (byName) return byName;
    }
    const body = upper(card.textContent);
    return plants.find((item) => [item?.denominazione, item?.nome, item?.idSap, item?.idSAP].map(text).filter(Boolean).some((value) => body.includes(upper(value)))) || null;
  }

  function getCoords(item) {
    if (!item) return null;
    try {
      if (typeof getImpiantoNavigationCoordinates === 'function') {
        const c = getImpiantoNavigationCoordinates(item);
        if (Number.isFinite(Number(c?.lat)) && Number.isFinite(Number(c?.lng))) return { lat: Number(c.lat), lng: Number(c.lng) };
      }
    } catch (_) {}
    try {
      const c = window.HeraCoordinateRepair?.getCoordinates?.(item);
      if (Number.isFinite(Number(c?.lat)) && Number.isFinite(Number(c?.lng))) return { lat: Number(c.lat), lng: Number(c.lng) };
    } catch (_) {}
    const lat = Number(item.lat ?? item.latitude ?? item.Latitudine ?? item.gpsY ?? item.GPSY);
    const lng = Number(item.lng ?? item.lon ?? item.longitude ?? item.Longitudine ?? item.gpsX ?? item.GPSX);
    return Number.isFinite(lat) && Number.isFinite(lng) ? { lat, lng } : null;
  }

  function isNavigate(el) {
    return /(^|\s)NAVIGA($|\s)/i.test(text(el?.textContent || el?.value || el?.getAttribute?.('aria-label')));
  }

  function isDone(el) {
    return /(^|\s)FATTO($|\s)/i.test(text(el?.textContent || el?.value || el?.getAttribute?.('aria-label')));
  }

  function findCard(nav) {
    let node = nav.parentElement;
    for (let depth = 0; node && depth < 8; depth += 1, node = node.parentElement) {
      if (node.matches?.('.map-popup-actions,.recommended-item')) continue;
      const actions = Array.from(node.querySelectorAll('button,a,[role="button"],input[type="button"],input[type="submit"]'));
      if (actions.some(isNavigate) && actions.some(isDone)) return node;
    }
    return null;
  }

  function ensureStyles() {
    if (document.getElementById('hera-street-view-card-style')) return;
    const style = document.createElement('style');
    style.id = 'hera-street-view-card-style';
    style.textContent = `
      .hera-street-view-anchor{position:relative!important}
      .hera-street-view-mini{position:absolute;left:8px;bottom:8px;width:44px;height:30px;min-width:44px;min-height:30px;padding:0;border:1.5px solid #9ca3af;border-radius:8px;background:#fff;box-shadow:0 2px 5px rgba(15,23,42,.14);display:flex;align-items:center;justify-content:center;font-size:17px;line-height:1;z-index:4;overflow:hidden}
      .hera-street-view-mini:active{transform:scale(.96)}
      .hera-street-view-mini[data-state="loading"]{opacity:.6;pointer-events:none}
      .hera-street-view-mini img{width:100%;height:100%;object-fit:cover;display:block}
      .hera-sv-modal{position:fixed;inset:0;z-index:2147483000;background:rgba(15,23,42,.72);display:flex;align-items:center;justify-content:center;padding:18px}
      .hera-sv-modal.hidden{display:none}
      .hera-sv-dialog{width:min(720px,100%);max-height:90vh;background:#fff;border-radius:18px;overflow:hidden;box-shadow:0 24px 60px rgba(0,0,0,.35)}
      .hera-sv-head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 14px;font-weight:800}
      .hera-sv-close{border:0;background:#eef2f7;border-radius:999px;width:36px;height:36px;font-size:20px}
      .hera-sv-body{padding:0 14px 14px}.hera-sv-body img{display:block;width:100%;max-height:65vh;object-fit:cover;border-radius:12px;background:#e5e7eb}
      .hera-sv-status{padding:22px;text-align:center;color:#475569;font-weight:700}
      @media(max-width:480px){.hera-street-view-mini{left:7px;bottom:7px;width:42px;height:28px;min-width:42px;min-height:28px;font-size:16px}}
    `;
    document.head.appendChild(style);
  }

  function ensureModal() {
    let modal = document.getElementById('hera-street-view-modal');
    if (modal) return modal;
    modal = document.createElement('div');
    modal.id = 'hera-street-view-modal';
    modal.className = 'hera-sv-modal hidden';
    modal.innerHTML = `<div class="hera-sv-dialog" role="dialog" aria-modal="true" aria-label="Street View impianto"><div class="hera-sv-head"><span>📸 Street View impianto</span><button type="button" class="hera-sv-close" aria-label="Chiudi">×</button></div><div class="hera-sv-body"><div class="hera-sv-status">Caricamento…</div></div></div>`;
    document.body.appendChild(modal);
    const close = () => modal.classList.add('hidden');
    modal.querySelector('.hera-sv-close')?.addEventListener('click', close);
    modal.addEventListener('click', (event) => { if (event.target === modal) close(); });
    return modal;
  }

  function showStatus(message) {
    const modal = ensureModal();
    modal.querySelector('.hera-sv-body').innerHTML = `<div class="hera-sv-status">${message}</div>`;
    modal.classList.remove('hidden');
  }

  function readCache() { try { return JSON.parse(localStorage.getItem(METADATA_CACHE_KEY) || '{}'); } catch (_) { return {}; } }
  function writeCache(cache) { try { localStorage.setItem(METADATA_CACHE_KEY, JSON.stringify(cache)); } catch (_) {} }
  function cacheKey(c) { return `${c.lat.toFixed(5)},${c.lng.toFixed(5)}`; }

  async function metadata(apiKey, coords) {
    const cache = readCache();
    const key = cacheKey(coords);
    const cached = cache[key];
    if (cached && Date.now() - Number(cached.at || 0) < METADATA_TTL) return cached.value;
    const url = new URL('https://maps.googleapis.com/maps/api/streetview/metadata');
    url.searchParams.set('location', `${coords.lat},${coords.lng}`);
    url.searchParams.set('source', 'outdoor');
    url.searchParams.set('key', apiKey);
    const response = await fetch(url);
    if (!response.ok) throw new Error(`Street View metadata ${response.status}`);
    const value = await response.json();
    cache[key] = { at: Date.now(), value };
    writeCache(cache);
    return value;
  }

  function imageUrl(apiKey, coords) {
    const url = new URL('https://maps.googleapis.com/maps/api/streetview');
    url.searchParams.set('size', '640x420');
    url.searchParams.set('location', `${coords.lat},${coords.lng}`);
    url.searchParams.set('fov', '90');
    url.searchParams.set('pitch', '0');
    url.searchParams.set('source', 'outdoor');
    url.searchParams.set('return_error_code', 'true');
    url.searchParams.set('key', apiKey);
    return url.toString();
  }

  async function openStreetView(nav, card, mini) {
    const apiKey = resolveApiKey();
    if (!apiKey) return showStatus('⚠️ Chiave Google API non disponibile.');
    const coords = getCoords(findPlant(nav, card));
    if (!coords) return showStatus('⚠️ Coordinate impianto non disponibili.');
    mini.dataset.state = 'loading'; mini.textContent = '…';
    try {
      const data = await metadata(apiKey, coords);
      if (data?.status !== 'OK') { mini.textContent = '📷'; return showStatus('Street View non disponibile vicino a questo impianto.'); }
      const c = Number.isFinite(Number(data?.location?.lat)) && Number.isFinite(Number(data?.location?.lng)) ? { lat: Number(data.location.lat), lng: Number(data.location.lng) } : coords;
      const src = imageUrl(apiKey, c);
      const modal = ensureModal();
      modal.querySelector('.hera-sv-body').innerHTML = `<img alt="Street View ingresso impianto" src="${src}">`;
      modal.classList.remove('hidden');
      mini.innerHTML = `<img alt="" aria-hidden="true" src="${src}">`;
    } catch (error) {
      console.warn('[STREET VIEW]', error); mini.textContent = '📷'; showStatus('⚠️ Street View non disponibile in questo momento.');
    } finally { delete mini.dataset.state; }
  }

  function enhance(nav) {
    if (!(nav instanceof HTMLElement) || nav.dataset.streetViewBound === '1' || !isNavigate(nav)) return;
    if (nav.closest('.map-popup-actions,.recommended-item')) return;
    const card = findCard(nav);
    if (!card || card.querySelector('.hera-street-view-mini')) return;
    card.classList.add('hera-street-view-anchor');
    const mini = document.createElement('button');
    mini.type = 'button'; mini.className = 'hera-street-view-mini'; mini.textContent = '📸';
    mini.title = 'Vedi ingresso con Street View'; mini.setAttribute('aria-label', 'Vedi ingresso con Street View');
    mini.addEventListener('click', (event) => { event.preventDefault(); event.stopPropagation(); openStreetView(nav, card, mini); });
    card.appendChild(mini); nav.dataset.streetViewBound = '1';
  }

  function scan(root = document) {
    root.querySelectorAll('button,a,[role="button"],input[type="button"],input[type="submit"]').forEach(enhance);
  }

  function scanVisibleCommessa() {
    const root = document.querySelector('#commessa-page:not(.hidden),#commessa-detail:not(.hidden),[data-page="commessa"]:not(.hidden)') || document;
    scan(root);
  }

  function install() {
    ensureStyles();
    ensureModal();
    // Nessun MutationObserver globale: scansione iniziale + eventi operativi già esistenti.
    requestAnimationFrame(scanVisibleCommessa);
    ['click','hera:commessa-rendered','hera:impianti-rendered','hera:recommended-traffic-weather'].forEach((name) => {
      window.addEventListener(name, () => setTimeout(scanVisibleCommessa, 80), { passive: true });
    });
    document.addEventListener('visibilitychange', () => { if (!document.hidden) setTimeout(scanVisibleCommessa, 80); });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true }); else install();
  window.HeraStreetViewCards = { installed: true, version: VERSION, refresh: scanVisibleCommessa, clearMetadataCache: () => { try { localStorage.removeItem(METADATA_CACHE_KEY); } catch (_) {} } };
})();
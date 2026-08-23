(() => {
  'use strict';
  if (window.HeraStreetViewCards?.installed && window.HeraStreetViewCards.version === '1.3.0') return;

  const VERSION = '1.3.0';
  const CACHE_KEY = 'heraStreetViewMetadataV1';
  const CACHE_TTL = 7 * 24 * 60 * 60 * 1000;
  let observer = null;

  const text = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();
  const upper = (value) => text(value).toLocaleUpperCase('it-IT');

  function resolveApiKey() {
    const candidates = [
      window.HERA_GOOGLE_MAPS_API_KEY,
      window.GOOGLE_MAPS_API_KEY,
      window.GOOGLE_API_KEY,
      window.mapsApiKey,
      window.googleMapsApiKey,
      window.__GOOGLE_MAPS_API_KEY__
    ].map(text).filter(Boolean);
    if (candidates.length) return candidates[0];
    try { return text(window.firebase?.app?.()?.options?.apiKey); } catch (_) {}
    try { return text(window.firebaseApp?.options?.apiKey); } catch (_) {}
    return '';
  }

  function getPlants() {
    try { if (typeof currentImpianti !== 'undefined' && Array.isArray(currentImpianti)) return currentImpianti; } catch (_) {}
    return Array.isArray(window.currentImpianti) ? window.currentImpianti : [];
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

  function findPlantFromRow(row) {
    const body = upper(row?.closest?.('.impianto-actions')?.parentElement?.textContent || row?.parentElement?.textContent || '');
    const plants = getPlants();
    return plants.find((item) => {
      const values = [item?.denominazione, item?.nome, item?.impianto, item?.idSap, item?.idSAP, item?.sap]
        .map(text).filter(Boolean);
      return values.some((value) => body.includes(upper(value)));
    }) || null;
  }

  function ensureStyles() {
    if (document.getElementById('hera-street-view-card-style')) return;
    const style = document.createElement('style');
    style.id = 'hera-street-view-card-style';
    style.textContent = `
      .impianto-primary-actions.hera-sv-row{position:relative!important}
      .hera-street-view-mini{position:absolute;left:8px;top:calc(100% + 5px);width:44px;height:30px;min-width:44px;min-height:30px;padding:0;border:1.5px solid #9ca3af;border-radius:8px;background:#fff;box-shadow:0 2px 5px rgba(15,23,42,.14);display:flex;align-items:center;justify-content:center;font-size:17px;line-height:1;z-index:30;overflow:hidden}
      .hera-street-view-mini:active{transform:scale(.96)}
      .hera-street-view-mini[data-state="loading"]{opacity:.6;pointer-events:none}
      .hera-street-view-mini img{width:100%;height:100%;object-fit:cover;display:block}
      .hera-sv-modal{position:fixed;inset:0;z-index:2147483000;background:rgba(15,23,42,.72);display:flex;align-items:center;justify-content:center;padding:18px}
      .hera-sv-modal.hidden{display:none}
      .hera-sv-dialog{width:min(720px,100%);max-height:90vh;background:#fff;border-radius:18px;overflow:hidden;box-shadow:0 24px 60px rgba(0,0,0,.35)}
      .hera-sv-head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 14px;font-weight:800}
      .hera-sv-close{border:0;background:#eef2f7;border-radius:999px;width:36px;height:36px;font-size:20px}
      .hera-sv-body{padding:0 14px 14px}
      .hera-sv-body img{display:block;width:100%;max-height:65vh;object-fit:cover;border-radius:12px;background:#e5e7eb}
      .hera-sv-status{padding:22px;text-align:center;color:#475569;font-weight:700}
      @media(max-width:480px){.hera-street-view-mini{left:7px;width:42px;height:28px;min-width:42px;min-height:28px;font-size:16px}}
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

  function readCache() { try { return JSON.parse(localStorage.getItem(CACHE_KEY) || '{}'); } catch (_) { return {}; } }
  function writeCache(cache) { try { localStorage.setItem(CACHE_KEY, JSON.stringify(cache)); } catch (_) {} }

  async function getMetadata(apiKey, coords) {
    const key = `${coords.lat.toFixed(5)},${coords.lng.toFixed(5)}`;
    const cache = readCache();
    const cached = cache[key];
    if (cached && Date.now() - Number(cached.at || 0) < CACHE_TTL) return cached.value;
    const url = new URL('https://maps.googleapis.com/maps/api/streetview/metadata');
    url.searchParams.set('location', `${coords.lat},${coords.lng}`);
    url.searchParams.set('source', 'outdoor');
    url.searchParams.set('key', apiKey);
    const response = await fetch(url.toString());
    if (!response.ok) throw new Error(`Street View metadata ${response.status}`);
    const value = await response.json();
    cache[key] = { at: Date.now(), value };
    writeCache(cache);
    return value;
  }

  function buildImageUrl(apiKey, coords) {
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

  async function openStreetView(row, mini) {
    const apiKey = resolveApiKey();
    if (!apiKey) return showStatus('⚠️ Chiave Google API non disponibile.');
    const plant = findPlantFromRow(row);
    const coords = getCoords(plant);
    if (!coords) return showStatus('⚠️ Coordinate impianto non disponibili.');

    mini.dataset.state = 'loading';
    mini.textContent = '…';
    try {
      const data = await getMetadata(apiKey, coords);
      if (data?.status !== 'OK') {
        mini.textContent = '📷';
        return showStatus('Street View non disponibile vicino a questo impianto.');
      }
      const c = Number.isFinite(Number(data?.location?.lat)) && Number.isFinite(Number(data?.location?.lng))
        ? { lat: Number(data.location.lat), lng: Number(data.location.lng) }
        : coords;
      const src = buildImageUrl(apiKey, c);
      const modal = ensureModal();
      modal.querySelector('.hera-sv-body').innerHTML = `<img alt="Street View ingresso impianto" src="${src}">`;
      modal.classList.remove('hidden');
      mini.innerHTML = `<img alt="" aria-hidden="true" src="${src}">`;
    } catch (error) {
      console.warn('[STREET VIEW]', error);
      mini.textContent = '📷';
      showStatus('⚠️ Street View non disponibile in questo momento.');
    } finally {
      delete mini.dataset.state;
    }
  }

  function enhanceRow(row) {
    if (!(row instanceof HTMLElement) || row.dataset.streetViewReady === '1') return;
    const nav = row.querySelector('[data-action-key="navigate"]');
    if (!nav) return;
    row.dataset.streetViewReady = '1';
    row.classList.add('hera-sv-row');
    const mini = document.createElement('button');
    mini.type = 'button';
    mini.className = 'hera-street-view-mini';
    mini.textContent = '📸';
    mini.title = 'Vedi ingresso con Street View';
    mini.setAttribute('aria-label', 'Vedi ingresso con Street View');
    mini.addEventListener('click', (event) => {
      event.preventDefault();
      event.stopPropagation();
      openStreetView(row, mini);
    });
    row.appendChild(mini);
  }

  function scan() {
    const list = document.getElementById('impianti-lista');
    if (!list) return false;
    list.querySelectorAll('.impianto-primary-actions').forEach(enhanceRow);
    if (!observer) {
      observer = new MutationObserver((mutations) => {
        for (const mutation of mutations) {
          for (const node of mutation.addedNodes) {
            if (!(node instanceof HTMLElement)) continue;
            if (node.matches?.('.impianto-primary-actions')) enhanceRow(node);
            node.querySelectorAll?.('.impianto-primary-actions').forEach(enhanceRow);
          }
        }
      });
      observer.observe(list, { childList: true, subtree: true });
    }
    return true;
  }

  function install() {
    ensureStyles();
    ensureModal();
    scan();
    let attempts = 0;
    const timer = setInterval(() => {
      attempts += 1;
      if (scan() || attempts >= 20) clearInterval(timer);
    }, 250);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true });
  else install();

  window.HeraStreetViewCards = {
    installed: true,
    version: VERSION,
    refresh: scan,
    clearMetadataCache: () => { try { localStorage.removeItem(CACHE_KEY); } catch (_) {} }
  };
})();
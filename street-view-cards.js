(() => {
  'use strict';
  if (window.HeraStreetViewCards?.installed && window.HeraStreetViewCards.version === '2.1.0') return;

  const VERSION = '2.1.0';
  const SEARCH_RADII = [50, 100, 250, 500, 1000];
  const MONTHLY_LIMIT = 4800;
  const USAGE_COLLECTION = 'systemUsage';
  let observer = null;
  let mapsLoaderPromise = null;
  let activePanorama = null;

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

  function resolveFirestore() {
    try { if (typeof db !== 'undefined' && db?.runTransaction) return db; } catch (_) {}
    try {
      if (window.firebase?.firestore && typeof window.firebase.firestore === 'function') {
        const firestore = window.firebase.firestore();
        if (firestore?.runTransaction) return firestore;
      }
    } catch (_) {}
    return null;
  }

  function getMonthKey(date = new Date()) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    return `${year}-${month}`;
  }

  function getCurrentUserInfo() {
    try {
      const user = window.firebase?.auth?.()?.currentUser;
      if (user) return { uid: user.uid || null, email: user.email || null };
    } catch (_) {}
    try {
      if (typeof currentUser !== 'undefined' && currentUser) {
        return { uid: currentUser.uid || null, email: currentUser.email || null };
      }
    } catch (_) {}
    return { uid: null, email: null };
  }

  async function reserveSharedMonthlySlot() {
    const firestore = resolveFirestore();
    if (!firestore) throw new Error('Contatore condiviso Firestore non disponibile');

    const monthKey = getMonthKey();
    const ref = firestore.collection(USAGE_COLLECTION).doc(`streetView360_${monthKey}`);
    const user = getCurrentUserInfo();

    return firestore.runTransaction(async (transaction) => {
      const snap = await transaction.get(ref);
      const data = snap.exists ? (snap.data() || {}) : {};
      const currentCount = Math.max(0, Number(data.count || 0));
      if (currentCount >= MONTHLY_LIMIT) {
        return { allowed: false, count: currentCount, limit: MONTHLY_LIMIT, monthKey };
      }

      const nextCount = currentCount + 1;
      const payload = {
        type: 'streetView360',
        month: monthKey,
        count: nextCount,
        limit: MONTHLY_LIMIT,
        updatedAt: window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || new Date().toISOString(),
        lastUserUid: user.uid || null,
        lastUserEmail: user.email || null
      };
      if (!snap.exists) payload.createdAt = window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || new Date().toISOString();
      transaction.set(ref, payload, { merge: true });
      return { allowed: true, count: nextCount, limit: MONTHLY_LIMIT, monthKey };
    });
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
    return getPlants().find((item) => {
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
      .hera-sv-modal{position:fixed;inset:0;z-index:2147483000;background:rgba(15,23,42,.78);display:flex;align-items:center;justify-content:center;padding:12px}
      .hera-sv-modal.hidden{display:none}
      .hera-sv-dialog{width:min(920px,100%);height:min(82vh,760px);background:#fff;border-radius:18px;overflow:hidden;box-shadow:0 24px 60px rgba(0,0,0,.35);display:flex;flex-direction:column}
      .hera-sv-head{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 14px;font-weight:800;flex:0 0 auto}
      .hera-sv-title{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
      .hera-sv-close{border:0;background:#eef2f7;border-radius:999px;width:40px;height:40px;min-width:40px;font-size:22px}
      .hera-sv-body{position:relative;flex:1 1 auto;min-height:0;background:#e5e7eb}
      .hera-sv-panorama{position:absolute;inset:0;width:100%;height:100%}
      .hera-sv-status{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;padding:24px;text-align:center;color:#475569;font-weight:700;background:#fff;z-index:2}
      .hera-sv-badge{position:absolute;left:12px;bottom:12px;z-index:3;background:rgba(15,23,42,.82);color:#fff;border-radius:999px;padding:7px 10px;font-size:12px;font-weight:800;pointer-events:none}
      @media(max-width:480px){.hera-street-view-mini{left:7px;width:42px;height:28px;min-width:42px;min-height:28px;font-size:16px}.hera-sv-modal{padding:6px}.hera-sv-dialog{height:78vh;border-radius:14px}}
    `;
    document.head.appendChild(style);
  }

  function ensureModal() {
    let modal = document.getElementById('hera-street-view-modal');
    if (modal) return modal;
    modal = document.createElement('div');
    modal.id = 'hera-street-view-modal';
    modal.className = 'hera-sv-modal hidden';
    modal.innerHTML = `<div class="hera-sv-dialog" role="dialog" aria-modal="true" aria-label="Street View 360 impianto"><div class="hera-sv-head"><span class="hera-sv-title">🌐 Street View 360° impianto</span><button type="button" class="hera-sv-close" aria-label="Chiudi">×</button></div><div class="hera-sv-body"><div class="hera-sv-status">Caricamento Street View 360°…</div></div></div>`;
    document.body.appendChild(modal);
    const close = () => {
      modal.classList.add('hidden');
      try { activePanorama?.setVisible(false); } catch (_) {}
      activePanorama = null;
    };
    modal.querySelector('.hera-sv-close')?.addEventListener('click', close);
    modal.addEventListener('click', (event) => { if (event.target === modal) close(); });
    return modal;
  }

  function showStatus(message) {
    const modal = ensureModal();
    modal.querySelector('.hera-sv-body').innerHTML = `<div class="hera-sv-status">${message}</div>`;
    modal.classList.remove('hidden');
  }

  function loadGoogleMaps(apiKey) {
    if (window.google?.maps?.StreetViewService && window.google?.maps?.StreetViewPanorama) return Promise.resolve(window.google.maps);
    if (mapsLoaderPromise) return mapsLoaderPromise;
    mapsLoaderPromise = new Promise((resolve, reject) => {
      const existing = Array.from(document.scripts).find((s) => /maps\.googleapis\.com\/maps\/api\/js/.test(s.src || ''));
      if (existing) {
        let attempts = 0;
        const timer = setInterval(() => {
          attempts += 1;
          if (window.google?.maps?.StreetViewService && window.google?.maps?.StreetViewPanorama) {
            clearInterval(timer);
            resolve(window.google.maps);
          } else if (attempts >= 60) {
            clearInterval(timer);
            reject(new Error('Maps JavaScript API non disponibile'));
          }
        }, 150);
        return;
      }
      const callbackName = `__heraStreetViewMapsReady_${Date.now()}`;
      const script = document.createElement('script');
      const url = new URL('https://maps.googleapis.com/maps/api/js');
      url.searchParams.set('key', apiKey);
      url.searchParams.set('v', 'weekly');
      url.searchParams.set('callback', callbackName);
      script.src = url.toString();
      script.async = true;
      script.defer = true;
      window[callbackName] = () => {
        try { delete window[callbackName]; } catch (_) {}
        if (window.google?.maps?.StreetViewService && window.google?.maps?.StreetViewPanorama) resolve(window.google.maps);
        else reject(new Error('Street View JavaScript non inizializzato'));
      };
      script.onerror = () => {
        try { delete window[callbackName]; } catch (_) {}
        reject(new Error('Caricamento Maps JavaScript API fallito'));
      };
      document.head.appendChild(script);
    }).catch((error) => {
      mapsLoaderPromise = null;
      throw error;
    });
    return mapsLoaderPromise;
  }

  function getPanoramaAtRadius(service, coords, radius) {
    return new Promise((resolve) => {
      service.getPanorama({
        location: coords,
        radius,
        source: window.google.maps.StreetViewSource.OUTDOOR,
        preference: window.google.maps.StreetViewPreference.NEAREST
      }, (data, status) => resolve({ data, status, radius }));
    });
  }

  async function findNearbyPanorama(service, coords) {
    for (const radius of SEARCH_RADII) {
      const result = await getPanoramaAtRadius(service, coords, radius);
      if (result.status === window.google.maps.StreetViewStatus.OK && result.data?.location?.pano) return result;
    }
    return null;
  }

  function computeHeading(from, to) {
    const lat1 = from.lat * Math.PI / 180;
    const lat2 = to.lat * Math.PI / 180;
    const dLng = (to.lng - from.lng) * Math.PI / 180;
    const y = Math.sin(dLng) * Math.cos(lat2);
    const x = Math.cos(lat1) * Math.sin(lat2) - Math.sin(lat1) * Math.cos(lat2) * Math.cos(dLng);
    return (Math.atan2(y, x) * 180 / Math.PI + 360) % 360;
  }

  async function openStreetView(row, mini) {
    const apiKey = resolveApiKey();
    if (!apiKey) return showStatus('⚠️ Chiave Google API non disponibile.');
    const plant = findPlantFromRow(row);
    const coords = getCoords(plant);
    if (!coords) return showStatus('⚠️ Coordinate impianto non disponibili.');

    mini.dataset.state = 'loading';
    mini.textContent = '…';
    showStatus('Caricamento Street View 360°…');

    try {
      const maps = await loadGoogleMaps(apiKey);
      const service = new maps.StreetViewService();
      const found = await findNearbyPanorama(service, coords);
      if (!found) {
        mini.textContent = '📷';
        return showStatus('Street View 360° non disponibile entro 1 km da questo impianto.');
      }

      const usage = await reserveSharedMonthlySlot();
      if (!usage.allowed) {
        mini.textContent = '⛔';
        return showStatus(`⛔ Limite Street View mensile raggiunto (${usage.count}/${usage.limit}). Il contatore riparte automaticamente il mese prossimo.`);
      }

      const panoLocation = found.data.location.latLng;
      const panoCoords = { lat: panoLocation.lat(), lng: panoLocation.lng() };
      const heading = computeHeading(panoCoords, coords);
      const modal = ensureModal();
      const body = modal.querySelector('.hera-sv-body');
      body.innerHTML = `<div class="hera-sv-panorama" aria-label="Street View 360 gradi"></div><div class="hera-sv-badge">360° · ${usage.count}/${usage.limit} questo mese · panorama entro ${found.radius} m</div>`;
      modal.classList.remove('hidden');

      activePanorama = new maps.StreetViewPanorama(body.querySelector('.hera-sv-panorama'), {
        pano: found.data.location.pano,
        pov: { heading, pitch: 0 },
        zoom: 1,
        visible: true,
        addressControl: true,
        linksControl: true,
        panControl: true,
        zoomControl: true,
        fullscreenControl: false,
        motionTracking: false,
        motionTrackingControl: false,
        clickToGo: true,
        scrollwheel: true,
        disableDefaultUI: false
      });
      mini.textContent = '🌐';
    } catch (error) {
      console.warn('[STREET VIEW 360]', error);
      mini.textContent = '📷';
      const message = text(error?.message);
      if (/contatore condiviso|permission|permesso|firestore/i.test(message)) {
        showStatus('⚠️ Street View bloccato per sicurezza: non riesco a verificare il contatore condiviso dei 4.800 utilizzi mensili.');
      } else {
        showStatus(`⚠️ Street View 360° non disponibile in questo momento. ${message}`);
      }
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
    mini.textContent = '🌐';
    mini.title = 'Apri Street View 360°';
    mini.setAttribute('aria-label', 'Apri Street View 360 gradi');
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
    monthlyLimit: MONTHLY_LIMIT,
    refresh: scan
  };
})();
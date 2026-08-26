(() => {
  'use strict';

  const VERSION = '2.3.3';
  if (window.HeraStreetViewCards?.installed && window.HeraStreetViewCards.version === VERSION) return;

  const SEARCH_RADII = [50, 100, 250, 500, 1000];
  const MONTHLY_LIMIT = 4800;
  const USAGE_COLLECTION = 'appConfig';
  let observer = null;
  let mapsLoaderPromise = null;
  let activePanorama = null;
  let activeRouteMap = null;
  let activeRouteRenderer = null;
  let activeRouteMarker = null;
  let activeRouteAnimationFrame = null;
  let activeRoutePolyline = null;
  let openRequestId = 0;

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

  function getCurrentUserInfo() {
    try {
      const user = window.firebase?.auth?.()?.currentUser;
      if (user) return { uid: user.uid || null, email: user.email || null };
    } catch (_) {}
    try {
      if (typeof currentUser !== 'undefined' && currentUser) return { uid: currentUser.uid || null, email: currentUser.email || null };
    } catch (_) {}
    return { uid: null, email: null };
  }

  function getMonthKey(date = new Date()) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  }

  async function reserveSharedMonthlySlot() {
    const firestore = resolveFirestore();
    if (!firestore) throw new Error('Contatore condiviso Firestore non disponibile');
    const monthKey = getMonthKey();
    const ref = firestore.collection(USAGE_COLLECTION).doc(`streetViewUsage_${monthKey}`);
    const user = getCurrentUserInfo();
    return firestore.runTransaction(async (transaction) => {
      const snap = await transaction.get(ref);
      const data = snap.exists ? (snap.data() || {}) : {};
      const currentCount = Math.max(0, Number(data.count || 0));
      if (currentCount >= MONTHLY_LIMIT) return { allowed: false, count: currentCount, limit: MONTHLY_LIMIT, monthKey };
      const nextCount = currentCount + 1;
      const now = window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || new Date().toISOString();
      const payload = { type: 'streetView360', month: monthKey, count: nextCount, limit: MONTHLY_LIMIT, updatedAt: now, lastUserUid: user.uid || null, lastUserEmail: user.email || null };
      if (!snap.exists) payload.createdAt = now;
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
        const lat = Number(c?.lat);
        const lng = Number(c?.lng ?? c?.lon);
        if (Number.isFinite(lat) && Number.isFinite(lng)) return { lat, lng };
      }
    } catch (_) {}
    const lat = Number(item.lat ?? item.latitude ?? item.Latitudine ?? item.gpsY ?? item.GPSY);
    const lng = Number(item.lng ?? item.lon ?? item.longitude ?? item.Longitudine ?? item.gpsX ?? item.GPSX);
    return Number.isFinite(lat) && Number.isFinite(lng) ? { lat, lng } : null;
  }

  function getPlantKey(item) {
    if (!item) return '';
    try {
      if (typeof buildImpiantoKey === 'function') return text(buildImpiantoKey(item));
    } catch (_) {}
    try {
      if (typeof window.buildImpiantoKey === 'function') return text(window.buildImpiantoKey(item));
    } catch (_) {}
    return text(item.id ?? item.impiantoId ?? item.idSap ?? item.idSAP ?? item.codiceHera);
  }

  function parseLatLngPair(value) {
    const raw = decodeURIComponent(text(value));
    const match = raw.match(/(-?\d{1,2}(?:\.\d+)?)\s*[,;]\s*(-?\d{1,3}(?:\.\d+)?)/);
    if (!match) return null;
    const lat = Number(match[1]);
    const lng = Number(match[2]);
    if (!Number.isFinite(lat) || !Number.isFinite(lng) || Math.abs(lat) > 90 || Math.abs(lng) > 180) return null;
    return { lat, lng };
  }

  function getCoordsFromNavigate(nav) {
    if (!(nav instanceof HTMLElement)) return null;
    const datasetPairs = [
      [nav.dataset?.lat, nav.dataset?.lng],
      [nav.dataset?.latitude, nav.dataset?.longitude],
      [nav.dataset?.destinationLat, nav.dataset?.destinationLng],
      [nav.dataset?.targetLat, nav.dataset?.targetLng]
    ];
    for (const [rawLat, rawLng] of datasetPairs) {
      const lat = Number(rawLat);
      const lng = Number(rawLng);
      if (Number.isFinite(lat) && Number.isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180) return { lat, lng };
    }

    const candidates = [nav.getAttribute('href'), nav.dataset?.href, nav.dataset?.url, nav.dataset?.destination, nav.dataset?.coords, nav.dataset?.coordinates].map(text).filter(Boolean);
    for (const candidate of candidates) {
      try {
        const url = new URL(candidate, window.location.href);
        for (const key of ['destination', 'query', 'q', 'daddr', 'll']) {
          const pair = parseLatLngPair(url.searchParams.get(key));
          if (pair) return pair;
        }
        const pathPair = parseLatLngPair(url.pathname);
        if (pathPair) return pathPair;
      } catch (_) {
        const pair = parseLatLngPair(candidate);
        if (pair) return pair;
      }
    }
    return null;
  }

  function findPlantFromRow(row) {
    const card = row?.closest?.('[data-impianto-key]');
    const cardKey = text(card?.dataset?.impiantoKey);
    if (cardKey) {
      return getPlants().find((item) => getPlantKey(item) === cardKey) || null;
    }

    const body = upper(row?.closest?.('.impianto-actions')?.parentElement?.textContent || row?.parentElement?.textContent || '');
    return getPlants().find((item) => {
      const values = [item?.denominazione, item?.nome, item?.impianto, item?.idSap, item?.idSAP, item?.sap].map(text).filter(Boolean);
      return values.some((value) => body.includes(upper(value)));
    }) || null;
  }

  function resolveRowCoords(row, nav) {
    const card = row?.closest?.('[data-impianto-key]');
    if (card?.dataset?.impiantoKey) {
      const exactPlant = findPlantFromRow(row);
      return getCoords(exactPlant);
    }
    const direct = getCoordsFromNavigate(nav);
    if (direct) return direct;
    return getCoords(findPlantFromRow(row));
  }

  function ensureStyles() {
    let style = document.getElementById('hera-street-view-card-style');
    if (!style) {
      style = document.createElement('style');
      style.id = 'hera-street-view-card-style';
      document.head.appendChild(style);
    }
    style.textContent = `
      .impianto-primary-actions.hera-sv-row{position:relative!important}
      .hera-street-view-mini{position:absolute;left:8px;top:calc(100% + 5px);width:44px;height:30px;min-width:44px;min-height:30px;padding:0;border:1.5px solid #9ca3af;border-radius:8px;background:#fff;box-shadow:0 2px 5px rgba(15,23,42,.14);display:flex;align-items:center;justify-content:center;font-size:17px;line-height:1;z-index:30}
      .hera-street-view-mini[data-state="loading"]{opacity:.6;pointer-events:none}
      .hera-sv-modal{position:fixed;inset:0;z-index:2147483000;background:#fff;display:block;padding:0}
      .hera-sv-modal.hidden{display:none}
      .hera-sv-dialog{position:absolute;inset:0;width:100vw;height:100dvh;max-width:none;max-height:none;background:#fff;border-radius:0;overflow:hidden;display:flex;flex-direction:column}
      .hera-sv-head{height:50px;box-sizing:border-box;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:7px 12px;font-weight:800;border-bottom:1px solid #dbe2ea;background:#fff;flex:0 0 50px;z-index:5}
      .hera-sv-title{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
      .hera-sv-close{border:0;background:#eef2f7;border-radius:999px;width:38px;height:38px;min-width:38px;font-size:21px}
      .hera-sv-body{position:relative;flex:1 1 auto;min-height:0;background:#e5e7eb}
      .hera-sv-layout{position:absolute;inset:0;display:grid;grid-template-rows:50% 50%;background:#fff}
      .hera-sv-panorama-wrap{position:relative;min-height:0;background:#e5e7eb;border-bottom:2px solid #fff}
      .hera-sv-panorama{position:absolute;inset:0;width:100%;height:100%}
      .hera-sv-badge{position:absolute;left:10px;bottom:10px;z-index:3;background:rgba(15,23,42,.82);color:#fff;border-radius:999px;padding:6px 9px;font-size:11px;font-weight:800;pointer-events:none}
      .hera-sv-status{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;padding:24px;text-align:center;color:#475569;font-weight:700;background:#fff;z-index:2}
      .hera-sv-route-panel{position:relative;min-height:0;background:#111827;overflow:hidden}
      .hera-sv-route-map{position:absolute;inset:0;width:100%;height:100%;background:#111827}
      .hera-sv-route-overlay{position:absolute;left:10px;right:10px;top:10px;z-index:4;display:flex;justify-content:space-between;gap:8px;pointer-events:none}
      .hera-sv-route-summary,.hera-sv-route-note{background:rgba(15,23,42,.86);color:#fff;border-radius:10px;padding:7px 9px;font-size:11px;font-weight:800;max-width:48%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
      .hera-sv-route-actions{position:absolute;left:10px;right:10px;bottom:10px;z-index:4;display:flex;gap:8px;align-items:center;pointer-events:none}
      .hera-sv-route-btn{pointer-events:auto;border:0;border-radius:10px;padding:8px 11px;font-weight:800;font-size:12px;white-space:nowrap;background:#fff;color:#111827;box-shadow:0 2px 10px rgba(0,0,0,.2)}
      .hera-sv-route-btn.primary{background:#111827;color:#fff}
      @media(max-width:480px){.hera-street-view-mini{left:7px;width:42px;height:28px;min-width:42px;min-height:28px;font-size:16px}.hera-sv-head{height:46px;flex-basis:46px;padding:5px 9px}.hera-sv-route-summary,.hera-sv-route-note{font-size:10px}.hera-sv-route-btn{padding:7px 9px;font-size:11px}}
    `;
  }

  function stopRouteAnimation() {
    if (activeRouteAnimationFrame) cancelAnimationFrame(activeRouteAnimationFrame);
    activeRouteAnimationFrame = null;
  }

  function clearRouteRuntime() {
    stopRouteAnimation();
    try { activeRouteRenderer?.setMap(null); } catch (_) {}
    try { activeRoutePolyline?.setMap(null); } catch (_) {}
    try { activeRouteMarker?.setMap(null); } catch (_) {}
    activeRouteRenderer = null;
    activeRoutePolyline = null;
    activeRouteMarker = null;
    activeRouteMap = null;
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
      openRequestId += 1;
      modal.classList.add('hidden');
      document.documentElement.style.overflow = '';
      document.body.style.overflow = '';
      try { activePanorama?.setVisible(false); } catch (_) {}
      activePanorama = null;
      clearRouteRuntime();
    };
    modal.querySelector('.hera-sv-close')?.addEventListener('click', close);
    return modal;
  }

  function showStatus(message) {
    const modal = ensureModal();
    clearRouteRuntime();
    modal.querySelector('.hera-sv-body').innerHTML = `<div class="hera-sv-status">${message}</div>`;
    modal.classList.remove('hidden');
    document.documentElement.style.overflow = 'hidden';
    document.body.style.overflow = 'hidden';
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
          if (window.google?.maps?.StreetViewService && window.google?.maps?.StreetViewPanorama) { clearInterval(timer); resolve(window.google.maps); }
          else if (attempts >= 60) { clearInterval(timer); reject(new Error('Maps JavaScript API non disponibile')); }
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
      script.onerror = () => reject(new Error('Caricamento Maps JavaScript API fallito'));
      document.head.appendChild(script);
    }).catch((error) => { mapsLoaderPromise = null; throw error; });
    return mapsLoaderPromise;
  }

  function getPanoramaAtRadius(service, coords, radius) {
    return new Promise((resolve) => {
      service.getPanorama({ location: coords, radius, source: window.google.maps.StreetViewSource.OUTDOOR, preference: window.google.maps.StreetViewPreference.NEAREST }, (data, status) => resolve({ data, status, radius }));
    });
  }

  async function findNearbyPanorama(service, coords) {
    for (const radius of SEARCH_RADII) {
      const result = await getPanoramaAtRadius(service, coords, radius);
      if ((String(result.status) === 'OK' || result.status === window.google.maps.StreetViewStatus.OK) && result.data?.location?.pano) return result;
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

  function distanceMeters(a, b) {
    const r = 6371000;
    const p1 = a.lat * Math.PI / 180;
    const p2 = b.lat * Math.PI / 180;
    const dp = (b.lat - a.lat) * Math.PI / 180;
    const dl = (b.lng - a.lng) * Math.PI / 180;
    const h = Math.sin(dp / 2) ** 2 + Math.cos(p1) * Math.cos(p2) * Math.sin(dl / 2) ** 2;
    return 2 * r * Math.atan2(Math.sqrt(h), Math.sqrt(1 - h));
  }

  function formatDistance(meters) {
    const value = Math.max(0, Number(meters || 0));
    return value < 1000 ? `${Math.round(value)} m` : `${(value / 1000).toFixed(value < 10000 ? 1 : 0)} km`;
  }

  function openGoogleMapsRoute(from, to, mode = 'driving') {
    const url = new URL('https://www.google.com/maps/dir/');
    url.searchParams.set('api', '1');
    url.searchParams.set('origin', `${from.lat},${from.lng}`);
    url.searchParams.set('destination', `${to.lat},${to.lng}`);
    url.searchParams.set('travelmode', mode);
    window.open(url.toString(), '_blank', 'noopener');
  }

  function requestDirections(maps, from, to, travelMode) {
    return new Promise((resolve) => {
      if (typeof maps.DirectionsService !== 'function') return resolve({ result: null, status: 'SERVICE_UNAVAILABLE' });
      try {
        const service = new maps.DirectionsService();
        service.route({ origin: from, destination: to, travelMode, provideRouteAlternatives: true }, (result, status) => {
          const ok = String(status) === 'OK' || status === maps.DirectionsStatus?.OK;
          resolve({ result: ok && result?.routes?.length ? result : null, status: String(status || 'UNKNOWN') });
        });
      } catch (error) {
        resolve({ result: null, status: text(error?.message) || 'ERROR' });
      }
    });
  }

  function chooseBestRoute(result) {
    const routes = Array.isArray(result?.routes) ? result.routes : [];
    if (!routes.length) return null;
    return routes.reduce((best, route) => {
      const d = Number(route?.legs?.[0]?.distance?.value || Number.MAX_SAFE_INTEGER);
      return !best || d < best.distance ? { route, distance: d } : best;
    }, null);
  }

  async function requestOsrmRoute(from, to) {
    const url = `https://router.project-osrm.org/route/v1/driving/${encodeURIComponent(from.lng)},${encodeURIComponent(from.lat)};${encodeURIComponent(to.lng)},${encodeURIComponent(to.lat)}?overview=full&geometries=geojson&steps=false`;
    try {
      const response = await fetch(url, { method: 'GET', mode: 'cors', cache: 'no-store' });
      if (!response.ok) return null;
      const json = await response.json();
      const route = json?.routes?.[0];
      const coords = route?.geometry?.coordinates;
      if (!Array.isArray(coords) || coords.length < 2) return null;
      return {
        path: coords.map(([lng, lat]) => ({ lat: Number(lat), lng: Number(lng) })).filter((p) => Number.isFinite(p.lat) && Number.isFinite(p.lng)),
        distance: Number(route.distance || 0),
        snappedStart: json?.waypoints?.[0]?.location ? { lng: Number(json.waypoints[0].location[0]), lat: Number(json.waypoints[0].location[1]) } : null,
        snappedEnd: json?.waypoints?.[1]?.location ? { lng: Number(json.waypoints[1].location[0]), lat: Number(json.waypoints[1].location[1]) } : null
      };
    } catch (error) {
      console.warn('[STREET VIEW ROUTE] OSRM non disponibile', error);
      return null;
    }
  }

  function fitRouteBounds(maps, path, from, to) {
    if (!activeRouteMap) return;
    const bounds = new maps.LatLngBounds();
    bounds.extend(from);
    bounds.extend(to);
    (path || []).forEach((p) => bounds.extend(p));
    activeRouteMap.fitBounds(bounds, 36);
  }

  function animateRouteMarker(maps, path, replayButton) {
    stopRouteAnimation();
    if (!activeRouteMap || !Array.isArray(path) || path.length < 2) return;
    try { activeRouteMarker?.setMap(null); } catch (_) {}
    activeRouteMarker = new maps.Marker({ map: activeRouteMap, position: path[0], title: 'Percorso verso impianto', label: { text: '➜', fontSize: '18px', fontWeight: '900' }, zIndex: 999 });
    if (replayButton) replayButton.disabled = true;
    const duration = Math.min(9000, Math.max(3200, path.length * 60));
    const startedAt = performance.now();
    const frame = (now) => {
      const progress = Math.min(1, (now - startedAt) / duration);
      const scaled = progress * (path.length - 1);
      const index = Math.min(path.length - 2, Math.floor(scaled));
      const fraction = scaled - index;
      const a = path[index], b = path[index + 1];
      const aLat = typeof a.lat === 'function' ? a.lat() : Number(a.lat);
      const aLng = typeof a.lng === 'function' ? a.lng() : Number(a.lng);
      const bLat = typeof b.lat === 'function' ? b.lat() : Number(b.lat);
      const bLng = typeof b.lng === 'function' ? b.lng() : Number(b.lng);
      activeRouteMarker?.setPosition({ lat: aLat + (bLat - aLat) * fraction, lng: aLng + (bLng - aLng) * fraction });
      if (progress < 1) activeRouteAnimationFrame = requestAnimationFrame(frame);
      else { activeRouteAnimationFrame = null; if (replayButton) replayButton.disabled = false; }
    };
    activeRouteAnimationFrame = requestAnimationFrame(frame);
  }

  async function renderAnimatedRoute(maps, container, from, to) {
    clearRouteRuntime();
    activeRouteMap = new maps.Map(container, {
      center: { lat: (from.lat + to.lat) / 2, lng: (from.lng + to.lng) / 2 },
      zoom: 18,
      mapTypeId: maps.MapTypeId?.SATELLITE || 'satellite',
      mapTypeControl: false,
      streetViewControl: false,
      fullscreenControl: false,
      gestureHandling: 'greedy'
    });

    new maps.Marker({ map: activeRouteMap, position: from, title: 'Punto panoramico', label: '📷' });
    new maps.Marker({ map: activeRouteMap, position: to, title: 'Impianto / cancello', label: '📍' });

    const panel = container.closest('.hera-sv-route-panel');
    const summary = panel?.querySelector('.hera-sv-route-summary');
    const replay = panel?.querySelector('[data-sv-route-replay]');
    const navigate = panel?.querySelector('[data-sv-route-navigate]');
    const note = panel?.querySelector('.hera-sv-route-note');

    let path = [];
    let routeDistance = distanceMeters(from, to);
    let modeUsed = 'driving';
    let routeLabel = '';

    const drive = await requestDirections(maps, from, to, maps.TravelMode?.DRIVING || 'DRIVING');
    const chosen = drive.result ? chooseBestRoute(drive.result) : null;

    if (chosen?.route) {
      const oneRouteResult = { ...drive.result, routes: [chosen.route] };
      activeRouteRenderer = new maps.DirectionsRenderer({ map: activeRouteMap, directions: oneRouteResult, suppressMarkers: true, preserveViewport: false, polylineOptions: { strokeWeight: 6, strokeOpacity: 0.92 } });
      path = chosen.route.overview_path || [];
      routeDistance = Number(chosen.route.legs?.[0]?.distance?.value || routeDistance);
      routeLabel = 'Percorso carrabile Google';
    } else {
      const osrm = await requestOsrmRoute(from, to);
      if (osrm?.path?.length > 1) {
        path = osrm.path;
        routeDistance = osrm.distance || routeDistance;
        activeRoutePolyline = new maps.Polyline({ map: activeRouteMap, path, strokeWeight: 6, strokeOpacity: 0.95, geodesic: false });
        if (osrm.snappedStart && distanceMeters(from, osrm.snappedStart) > 4) {
          new maps.Polyline({ map: activeRouteMap, path: [from, osrm.snappedStart], strokeWeight: 3, strokeOpacity: 0.65, geodesic: true });
        }
        if (osrm.snappedEnd && distanceMeters(osrm.snappedEnd, to) > 4) {
          new maps.Polyline({ map: activeRouteMap, path: [osrm.snappedEnd, to], strokeWeight: 3, strokeOpacity: 0.65, geodesic: true });
        }
        fitRouteBounds(maps, path, from, to);
        routeLabel = 'Percorso carrabile stradale';
      }
    }

    if (!path.length) {
      fitRouteBounds(maps, [], from, to);
      routeLabel = 'Percorso stradale non disponibile';
    }

    if (summary) summary.textContent = `📷 Panorama → 📍 Impianto · ${formatDistance(routeDistance)}`;
    if (note) note.textContent = routeLabel;
    navigate?.addEventListener('click', () => openGoogleMapsRoute(from, to, modeUsed));
    if (path.length > 1) {
      replay?.addEventListener('click', () => animateRouteMarker(maps, path, replay));
      window.setTimeout(() => animateRouteMarker(maps, path, replay), 350);
    } else if (replay) {
      replay.disabled = true;
      replay.textContent = 'PERCORSO NON DISPONIBILE';
    }
  }

  async function openStreetView(row, mini, nav) {
    const requestId = ++openRequestId;
    const apiKey = resolveApiKey();
    if (!apiKey) return showStatus('⚠️ Chiave Google API non disponibile.');
    const coords = resolveRowCoords(row, nav);
    if (!coords) return showStatus('⚠️ Coordinate impianto non disponibili per la scheda selezionata.');

    mini.dataset.state = 'loading';
    mini.textContent = '…';
    showStatus('Caricamento Street View 360°…');

    try {
      const maps = await loadGoogleMaps(apiKey);
      if (requestId !== openRequestId) return;
      const service = new maps.StreetViewService();
      const found = await findNearbyPanorama(service, coords);
      if (requestId !== openRequestId) return;
      if (!found) { mini.textContent = '📷'; return showStatus('Street View 360° non disponibile entro 1 km da questo impianto.'); }

      const usage = await reserveSharedMonthlySlot();
      if (requestId !== openRequestId) return;
      if (!usage.allowed) { mini.textContent = '⛔'; return showStatus(`⛔ Limite Street View mensile raggiunto (${usage.count}/${usage.limit}).`); }

      const panoLocation = found.data.location.latLng;
      const panoCoords = { lat: panoLocation.lat(), lng: panoLocation.lng() };
      const heading = computeHeading(panoCoords, coords);
      const directDistance = distanceMeters(panoCoords, coords);
      const modal = ensureModal();
      const body = modal.querySelector('.hera-sv-body');
      body.innerHTML = `<div class="hera-sv-layout"><div class="hera-sv-panorama-wrap"><div class="hera-sv-panorama" aria-label="Street View 360 gradi"></div><div class="hera-sv-badge">360° · ${usage.count}/${usage.limit} questo mese · panorama a circa ${formatDistance(directDistance)}</div></div><div class="hera-sv-route-panel"><div class="hera-sv-route-map" aria-label="Percorso animato dalla panoramica all'impianto"></div><div class="hera-sv-route-overlay"><span class="hera-sv-route-summary">📷 Panorama → 📍 Impianto · calcolo percorso…</span><span class="hera-sv-route-note">Calcolo percorso stradale…</span></div><div class="hera-sv-route-actions"><button type="button" class="hera-sv-route-btn primary" data-sv-route-replay>↻ RIPETI PERCORSO</button><button type="button" class="hera-sv-route-btn" data-sv-route-navigate>🧭 APRI IN MAPS</button></div></div></div>`;
      modal.classList.remove('hidden');
      document.documentElement.style.overflow = 'hidden';
      document.body.style.overflow = 'hidden';

      activePanorama = new maps.StreetViewPanorama(body.querySelector('.hera-sv-panorama'), {
        pano: found.data.location.pano,
        pov: { heading, pitch: 0 }, zoom: 1, visible: true,
        addressControl: true, linksControl: true, panControl: true, zoomControl: true,
        fullscreenControl: false, motionTracking: false, motionTrackingControl: false,
        clickToGo: true, scrollwheel: true, disableDefaultUI: false
      });

      await renderAnimatedRoute(maps, body.querySelector('.hera-sv-route-map'), panoCoords, coords);
      if (requestId !== openRequestId) return;
      mini.textContent = '🌐';
    } catch (error) {
      if (requestId !== openRequestId) return;
      console.warn('[STREET VIEW 360]', error);
      mini.textContent = '📷';
      const message = text(error?.message);
      if (/contatore condiviso|permission|permesso|firestore/i.test(message)) showStatus('⚠️ Street View bloccato: non riesco a verificare il contatore condiviso.');
      else showStatus(`⚠️ Street View 360° non disponibile in questo momento. ${message}`);
    } finally {
      if (requestId === openRequestId) delete mini.dataset.state;
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
    mini.addEventListener('click', (event) => { event.preventDefault(); event.stopPropagation(); openStreetView(row, mini, nav); });
    row.appendChild(mini);
  }

  function scan() {
    const list = document.getElementById('impianti-lista');
    if (!list) return false;
    list.querySelectorAll('.impianto-primary-actions').forEach(enhanceRow);
    if (!observer) {
      observer = new MutationObserver((mutations) => {
        for (const mutation of mutations) for (const node of mutation.addedNodes) {
          if (!(node instanceof HTMLElement)) continue;
          if (node.matches?.('.impianto-primary-actions')) enhanceRow(node);
          node.querySelectorAll?.('.impianto-primary-actions').forEach(enhanceRow);
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
    const timer = setInterval(() => { attempts += 1; if (scan() || attempts >= 20) clearInterval(timer); }, 250);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true });
  else install();

  window.HeraStreetViewCards = { installed: true, version: VERSION, monthlyLimit: MONTHLY_LIMIT, refresh: scan };
})();
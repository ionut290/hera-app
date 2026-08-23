(() => {
  'use strict';
  if (window.HeraRecommendedTrafficWeather?.installed) return;

  const VERSION = '1.0.0';
  const ROUTE_TTL = 10 * 60 * 1000;
  const WEATHER_TTL = 15 * 60 * 1000;
  const MAX_ROUTE_REQUESTS = 10;
  const MAX_WEATHER_ZONES = 6;
  const ROUTE_CACHE_KEY = 'heraTrafficRouteCacheV1';
  const WEATHER_CACHE_KEY = 'heraTrafficWeatherCacheV1';
  let running = false;
  let rerunTimer = 0;

  const text = v => String(v ?? '').trim();
  const clamp = (v, min, max) => Math.min(max, Math.max(min, v));
  const parseDuration = value => {
    const match = text(value).match(/^([0-9.]+)s$/);
    return match ? Number(match[1]) / 60 : 0;
  };
  const readJson = (key) => {
    try { return JSON.parse(localStorage.getItem(key) || '{}'); } catch (_) { return {}; }
  };
  const writeJson = (key, value) => {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch (_) {}
  };
  const prune = (cache, ttl) => {
    const now = Date.now();
    Object.keys(cache).forEach(k => { if (!cache[k]?.at || now - cache[k].at > ttl) delete cache[k]; });
    return cache;
  };

  function resolveApiKey() {
    const globals = [
      window.HERA_GOOGLE_MAPS_API_KEY,
      window.GOOGLE_MAPS_API_KEY,
      window.GOOGLE_API_KEY,
      window.mapsApiKey,
      window.googleMapsApiKey,
      window.__GOOGLE_MAPS_API_KEY__
    ].map(text).filter(Boolean);
    if (globals.length) return globals[0];
    try {
      const key = text(window.firebase?.app?.()?.options?.apiKey);
      if (key) return key;
    } catch (_) {}
    try {
      const key = text(window.firebaseApp?.options?.apiKey);
      if (key) return key;
    } catch (_) {}
    return '';
  }

  function roundedZone(coords) {
    return `${(Math.round(coords.lat * 20) / 20).toFixed(2)},${(Math.round(coords.lng * 20) / 20).toFixed(2)}`;
  }

  function routeCacheKey(a, b) {
    const r = n => Number(n).toFixed(4);
    return `${r(a.lat)},${r(a.lng)}>${r(b.lat)},${r(b.lng)}`;
  }

  async function trafficRoute(apiKey, origin, destination) {
    const cache = prune(readJson(ROUTE_CACHE_KEY), ROUTE_TTL);
    const key = routeCacheKey(origin, destination);
    if (cache[key]) return cache[key].value;

    const response = await fetch('https://routes.googleapis.com/directions/v2:computeRoutes', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-Goog-Api-Key': apiKey,
        'X-Goog-FieldMask': 'routes.duration,routes.staticDuration,routes.distanceMeters'
      },
      body: JSON.stringify({
        origin: { location: { latLng: { latitude: origin.lat, longitude: origin.lng } } },
        destination: { location: { latLng: { latitude: destination.lat, longitude: destination.lng } } },
        travelMode: 'DRIVE',
        routingPreference: 'TRAFFIC_AWARE'
      })
    });
    if (!response.ok) throw new Error(`Routes API ${response.status}`);
    const json = await response.json();
    const route = json?.routes?.[0];
    if (!route) throw new Error('Routes API senza percorso');
    const value = {
      durationMinutes: Math.max(1, Math.round(parseDuration(route.duration))),
      staticMinutes: Math.max(1, Math.round(parseDuration(route.staticDuration || route.duration))),
      distanceKm: Number(route.distanceMeters || 0) / 1000
    };
    value.delayMinutes = Math.max(0, value.durationMinutes - value.staticMinutes);
    value.delayPct = value.staticMinutes ? Math.round(value.delayMinutes / value.staticMinutes * 100) : 0;
    cache[key] = { at: Date.now(), value };
    writeJson(ROUTE_CACHE_KEY, cache);
    return value;
  }

  function weatherRiskFromHours(hours) {
    let factor = 1;
    let level = 'ok';
    let label = 'Meteo regolare';
    let precip = 0, thunder = 0, gust = 0, temp = -99;
    for (const h of hours || []) {
      const p = Number(h?.precipitation?.probability?.percent ?? h?.precipitationProbability?.percent ?? h?.precipitationProbability ?? 0) || 0;
      const t = Number(h?.thunderstormProbability ?? h?.thunderstormProbability?.percent ?? 0) || 0;
      const g = Number(h?.wind?.gust?.value ?? h?.windGust?.value ?? h?.windGust ?? 0) || 0;
      const c = Number(h?.temperature?.degrees ?? h?.temperature?.value ?? 0) || 0;
      precip = Math.max(precip, p); thunder = Math.max(thunder, t); gust = Math.max(gust, g); temp = Math.max(temp, c);
    }
    if (precip >= 70) { factor += 0.18; level = 'warning'; label = 'Pioggia probabile'; }
    else if (precip >= 40) { factor += 0.08; level = 'attention'; label = 'Possibile pioggia'; }
    if (thunder >= 40) { factor += 0.20; level = 'warning'; label = 'Rischio temporali'; }
    if (gust >= 70) { factor += 0.30; level = 'warning'; label = 'Raffiche forti'; }
    else if (gust >= 50) { factor += 0.12; if (level === 'ok') level = 'attention'; label = 'Vento sostenuto'; }
    if (temp >= 35) { factor += 0.10; if (level === 'ok') level = 'attention'; label = 'Caldo intenso'; }
    factor = clamp(factor, 1, 1.5);
    return { factor, level, label, precip, thunder, gust, temp };
  }

  async function weatherFor(apiKey, coords) {
    const cache = prune(readJson(WEATHER_CACHE_KEY), WEATHER_TTL);
    const zone = roundedZone(coords);
    if (cache[zone]) return cache[zone].value;
    const url = new URL('https://weather.googleapis.com/v1/forecast/hours:lookup');
    url.searchParams.set('key', apiKey);
    url.searchParams.set('location.latitude', coords.lat);
    url.searchParams.set('location.longitude', coords.lng);
    url.searchParams.set('hours', '4');
    url.searchParams.set('pageSize', '4');
    url.searchParams.set('unitsSystem', 'METRIC');
    const response = await fetch(url.toString());
    if (!response.ok) throw new Error(`Weather API ${response.status}`);
    const json = await response.json();
    const value = weatherRiskFromHours(json?.forecastHours || []);
    cache[zone] = { at: Date.now(), value };
    writeJson(WEATHER_CACHE_KEY, cache);
    return value;
  }

  function currentPosition(timeout = 2500) {
    return new Promise(resolve => {
      if (!navigator.geolocation) return resolve(null);
      navigator.geolocation.getCurrentPosition(
        p => resolve({ lat: p.coords.latitude, lng: p.coords.longitude }),
        () => resolve(null),
        { enableHighAccuracy: false, timeout, maximumAge: 120000 }
      );
    });
  }

  function mutateTextNode(element, value) {
    if (!element) return;
    const node = [...element.childNodes].find(n => n.nodeType === Node.TEXT_NODE);
    if (node) node.nodeValue = value;
  }

  function annotatePanel(enriched) {
    const panel = document.getElementById('recommended-plants-panel');
    if (!panel || panel.classList.contains('hidden')) return;
    const routeById = new Map(enriched.map(e => [text(e.item?.physicalPlantId || e.item?.impiantoId || e.item?.id || e.item?.idSap || e.item?.denominazione || e.item?.nome), e]));
    panel.querySelectorAll('[data-recommended-id]').forEach(card => {
      const entry = routeById.get(text(card.dataset.recommendedId));
      if (!entry) return;
      const smalls = card.querySelectorAll('.recommended-main small');
      if (smalls[0]) {
        const weather = entry.weather;
        const traffic = entry.traffic;
        const suffix = `${traffic ? ` · 🚦 ${traffic.durationMinutes} min${traffic.delayMinutes ? ` (+${traffic.delayMinutes})` : ''}` : ''}${weather ? ` · 🌦️ ${weather.label}${weather.factor > 1 ? ` +${Math.round((weather.factor - 1) * 100)}%` : ''}` : ''}`;
        const firstText = [...smalls[0].childNodes].find(n => n.nodeType === Node.TEXT_NODE);
        if (firstText && !text(firstText.nodeValue).includes('🚦')) firstText.nodeValue = `${firstText.nodeValue}${suffix}`;
      }
    });
    const summarySpans = panel.querySelectorAll('.recommended-summary span');
    const trafficTotal = enriched.reduce((s, e) => s + Number(e.traffic?.delayMinutes || 0), 0);
    const weatherAffected = enriched.filter(e => e.weather?.factor > 1).length;
    if (summarySpans[3]) mutateTextNode(summarySpans[3], `🚦 traffico +${Math.round(trafficTotal)} min`);
    if (summarySpans[4] && weatherAffected) mutateTextNode(summarySpans[4], `🌦️ ${weatherAffected} meteo da considerare`);
  }

  async function enrich() {
    if (running) return;
    const panel = document.getElementById('recommended-plants-panel');
    if (!panel || panel.classList.contains('hidden')) return;
    const api = window.HeraRecommendedPlants;
    const state = api?.getState?.();
    const plan = state?.lastPlan;
    if (!Array.isArray(plan) || !plan.length) return;
    const apiKey = resolveApiKey();
    if (!apiKey) {
      console.warn('Traffico/meteo: chiave Google API non trovata.');
      return;
    }

    running = true;
    try {
      const configStart = api?.config?.start || { lat: 44.5790, lng: 11.3635 };
      let origin = (!state.originMode || state.originMode === 'auto' || state.originMode === 'avola') ? { lat: configStart.lat, lng: configStart.lng } : await currentPosition();
      if (!origin) origin = { lat: configStart.lat, lng: configStart.lng };
      const candidates = plan.slice(0, MAX_ROUTE_REQUESTS);
      const zones = new Map();
      candidates.forEach(entry => {
        if (entry?.coords && zones.size < MAX_WEATHER_ZONES) zones.set(roundedZone(entry.coords), entry.coords);
      });
      const weatherResults = new Map();
      await Promise.all([...zones.entries()].map(async ([zone, coords]) => {
        try { weatherResults.set(zone, await weatherFor(apiKey, coords)); } catch (error) { console.warn('Weather API non disponibile:', error); }
      }));

      let cumulative = 0;
      const enriched = [];
      for (const entry of candidates) {
        if (!entry?.coords) continue;
        let traffic = null;
        try { traffic = await trafficRoute(apiKey, origin, entry.coords); } catch (error) { console.warn('Routes API traffico non disponibile:', error); }
        const weather = weatherResults.get(roundedZone(entry.coords)) || null;
        if (traffic?.durationMinutes) {
          entry.traffic = traffic;
          entry.driveMinutes = traffic.durationMinutes;
          if (traffic.distanceKm > 0) entry.km = traffic.distanceKm;
        }
        if (weather) {
          entry.weather = weather;
          if (!entry.__weatherBaseWorkMinutes) entry.__weatherBaseWorkMinutes = entry.workMinutes;
          entry.workMinutes = Math.max(1, Math.round(entry.__weatherBaseWorkMinutes * weather.factor));
        }
        cumulative += Number(entry.driveMinutes || 0) + Number(entry.workMinutes || 0) + Number(entry.unload?.driveMinutes || 0) + Number(entry.unload?.unloadMinutes || 0);
        entry.cumulativeMinutes = cumulative;
        entry.fitsDay = cumulative <= Number(api?.config?.planningMinutes || 480);
        enriched.push(entry);
        origin = entry.unload ? { lat: configStart.lat, lng: configStart.lng } : entry.coords;
      }

      window.HeraAdaptiveWorkLearning?.applyToRecommendedPanel?.();
      window.setTimeout(() => annotatePanel(enriched), 180);
      window.dispatchEvent(new CustomEvent('hera:recommended-traffic-weather', { detail: { count: enriched.length, at: Date.now() } }));
    } finally {
      running = false;
    }
  }

  function schedule() {
    clearTimeout(rerunTimer);
    rerunTimer = window.setTimeout(enrich, 250);
  }

  document.addEventListener('click', event => {
    const button = event.target?.closest?.('#recommended-plants-btn,#recommended-origin-avola,#recommended-origin-live');
    if (button) window.setTimeout(schedule, 500);
  }, true);
  window.addEventListener('hera:adaptive-learning-sample', schedule);
  window.addEventListener('online', schedule);

  window.HeraRecommendedTrafficWeather = {
    installed: true,
    version: VERSION,
    refresh: enrich,
    clearCache: () => {
      try { localStorage.removeItem(ROUTE_CACHE_KEY); localStorage.removeItem(WEATHER_CACHE_KEY); } catch (_) {}
    }
  };
})();

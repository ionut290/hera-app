(() => {
  'use strict';

  const VERSION = '1.1.0-responsive2';
  if (window.HeraRecommendedTrafficWeather?.version === VERSION) return;

  const ROUTE_TTL = 10 * 60 * 1000;
  const WEATHER_TTL = 15 * 60 * 1000;
  const MAX_ROUTE_REQUESTS = 8;
  const MAX_WEATHER_ZONES = 6;
  const ROUTE_CACHE_KEY = 'heraTrafficRouteCacheV1';
  const WEATHER_CACHE_KEY = 'heraTrafficWeatherCacheV1';

  let running = false;
  let rerunRequested = false;
  let scheduleTimer = 0;
  let idleTask = 0;
  let runSequence = 0;

  const text = (value) => String(value ?? '').trim();
  const clamp = (value, min, max) => Math.min(max, Math.max(min, value));

  function readJson(key) {
    try {
      return JSON.parse(localStorage.getItem(key) || '{}');
    } catch (_) {
      return {};
    }
  }

  function writeJson(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch (_) {}
  }

  function prune(cache, ttl) {
    const timestamp = Date.now();
    Object.keys(cache).forEach((key) => {
      if (!cache[key]?.at || timestamp - cache[key].at > ttl) delete cache[key];
    });
    return cache;
  }

  function parseDuration(value) {
    const match = text(value).match(/^([0-9.]+)s$/);
    return match ? Number(match[1]) / 60 : 0;
  }

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

  function routeCacheKey(origin, destination) {
    const rounded = (value) => Number(value).toFixed(4);
    return `${rounded(origin.lat)},${rounded(origin.lng)}>${rounded(destination.lat)},${rounded(destination.lng)}`;
  }

  function panelIsOpen() {
    const panel = document.getElementById('recommended-plants-panel');
    return Boolean(panel && !panel.classList.contains('hidden'));
  }

  async function trafficRoute(apiKey, origin, destination, cache) {
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
    value.delayPct = value.staticMinutes
      ? Math.round(value.delayMinutes / value.staticMinutes * 100)
      : 0;
    cache[key] = { at: Date.now(), value };
    return value;
  }

  function weatherRiskFromHours(hours) {
    let factor = 1;
    let level = 'ok';
    let label = 'Meteo regolare';
    let precip = 0;
    let thunder = 0;
    let gust = 0;
    let temp = -99;

    for (const hour of hours || []) {
      const precipitation = Number(
        hour?.precipitation?.probability?.percent
        ?? hour?.precipitationProbability?.percent
        ?? hour?.precipitationProbability
        ?? 0
      ) || 0;
      const thunderstorm = Number(
        hour?.thunderstormProbability?.percent
        ?? hour?.thunderstormProbability
        ?? 0
      ) || 0;
      const windGust = Number(
        hour?.wind?.gust?.value
        ?? hour?.windGust?.value
        ?? hour?.windGust
        ?? 0
      ) || 0;
      const temperature = Number(
        hour?.temperature?.degrees
        ?? hour?.temperature?.value
        ?? 0
      ) || 0;
      precip = Math.max(precip, precipitation);
      thunder = Math.max(thunder, thunderstorm);
      gust = Math.max(gust, windGust);
      temp = Math.max(temp, temperature);
    }

    if (precip >= 70) {
      factor += 0.18;
      level = 'warning';
      label = 'Pioggia probabile';
    } else if (precip >= 40) {
      factor += 0.08;
      level = 'attention';
      label = 'Possibile pioggia';
    }
    if (thunder >= 40) {
      factor += 0.20;
      level = 'warning';
      label = 'Rischio temporali';
    }
    if (gust >= 70) {
      factor += 0.30;
      level = 'warning';
      label = 'Raffiche forti';
    } else if (gust >= 50) {
      factor += 0.12;
      if (level === 'ok') level = 'attention';
      label = 'Vento sostenuto';
    }
    if (temp >= 35) {
      factor += 0.10;
      if (level === 'ok') level = 'attention';
      label = 'Caldo intenso';
    }

    return {
      factor: clamp(factor, 1, 1.5),
      level,
      label,
      precip,
      thunder,
      gust,
      temp
    };
  }

  async function weatherFor(apiKey, coords, cache) {
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
    return value;
  }

  function currentPosition(timeout = 2200) {
    return new Promise((resolve) => {
      if (!navigator.geolocation) return resolve(null);
      navigator.geolocation.getCurrentPosition(
        (position) => resolve({ lat: position.coords.latitude, lng: position.coords.longitude }),
        () => resolve(null),
        { enableHighAccuracy: false, timeout, maximumAge: 120000 }
      );
    });
  }

  function cancelScheduled() {
    window.clearTimeout(scheduleTimer);
    scheduleTimer = 0;
    if (idleTask && typeof window.cancelIdleCallback === 'function') {
      window.cancelIdleCallback(idleTask);
    }
    idleTask = 0;
  }

  function cancel() {
    cancelScheduled();
    runSequence += 1;
    rerunRequested = false;
  }

  async function enrich() {
    scheduleTimer = 0;
    idleTask = 0;
    if (!panelIsOpen()) return;
    if (running) {
      rerunRequested = true;
      return;
    }

    const api = window.HeraRecommendedPlants;
    const recommendedState = api?.getState?.();
    const plan = recommendedState?.lastPlan;
    if (!Array.isArray(plan) || !plan.length) return;

    const apiKey = resolveApiKey();
    if (!apiKey) {
      console.warn('Traffico/meteo: chiave Google API non trovata.');
      return;
    }

    running = true;
    const sequence = ++runSequence;
    const isCurrent = () => sequence === runSequence && panelIsOpen();
    const routeCache = prune(readJson(ROUTE_CACHE_KEY), ROUTE_TTL);
    const weatherCache = prune(readJson(WEATHER_CACHE_KEY), WEATHER_TTL);
    let routeCacheChanged = false;
    let weatherCacheChanged = false;

    try {
      const configStart = api?.config?.start || { lat: 44.579, lng: 11.3635 };
      let origin = (!recommendedState.originMode
        || recommendedState.originMode === 'auto'
        || recommendedState.originMode === 'avola')
        ? { lat: configStart.lat, lng: configStart.lng }
        : await currentPosition();
      if (!isCurrent()) return;
      if (!origin) origin = { lat: configStart.lat, lng: configStart.lng };

      const candidates = plan.slice(0, MAX_ROUTE_REQUESTS);
      const zones = new Map();
      for (const entry of candidates) {
        if (entry?.coords && zones.size < MAX_WEATHER_ZONES) {
          zones.set(roundedZone(entry.coords), entry.coords);
        }
      }

      const weatherResults = new Map();
      await Promise.all([...zones.entries()].map(async ([zone, coords]) => {
        const existed = Boolean(weatherCache[zone]);
        try {
          const value = await weatherFor(apiKey, coords, weatherCache);
          weatherResults.set(zone, value);
          if (!existed && weatherCache[zone]) weatherCacheChanged = true;
        } catch (error) {
          console.warn('Weather API non disponibile:', error);
        }
      }));
      if (!isCurrent()) return;

      let cumulative = 0;
      let enrichedCount = 0;
      for (const entry of candidates) {
        if (!entry?.coords || !isCurrent()) break;
        const cacheKey = routeCacheKey(origin, entry.coords);
        const existed = Boolean(routeCache[cacheKey]);
        let traffic = null;
        try {
          traffic = await trafficRoute(apiKey, origin, entry.coords, routeCache);
          if (!existed && routeCache[cacheKey]) routeCacheChanged = true;
        } catch (error) {
          console.warn('Routes API traffico non disponibile:', error);
        }
        if (!isCurrent()) break;

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

        cumulative += Number(entry.driveMinutes || 0)
          + Number(entry.workMinutes || 0)
          + Number(entry.unload?.driveMinutes || 0)
          + Number(entry.unload?.unloadMinutes || 0);
        entry.cumulativeMinutes = cumulative;
        entry.fitsDay = cumulative <= Number(api?.config?.planningMinutes || 480);
        enrichedCount += 1;
        origin = entry.unload
          ? { lat: configStart.lat, lng: configStart.lng }
          : entry.coords;
      }

      if (!isCurrent()) return;
      if (routeCacheChanged) writeJson(ROUTE_CACHE_KEY, routeCache);
      if (weatherCacheChanged) writeJson(WEATHER_CACHE_KEY, weatherCache);
      api?.refreshDecorations?.();
      window.dispatchEvent(new CustomEvent('hera:recommended-traffic-weather', {
        detail: { count: enrichedCount, at: Date.now() }
      }));
    } finally {
      running = false;
      if (rerunRequested && panelIsOpen()) {
        rerunRequested = false;
        schedule(500);
      }
    }
  }

  function schedule(delay = 650) {
    cancelScheduled();
    scheduleTimer = window.setTimeout(() => {
      scheduleTimer = 0;
      if (!panelIsOpen()) return;
      if (typeof window.requestIdleCallback === 'function') {
        idleTask = window.requestIdleCallback(enrich, { timeout: 1800 });
      } else {
        enrich();
      }
    }, delay);
  }

  window.addEventListener('hera:recommended-ready', () => schedule(650));
  window.addEventListener('hera:recommended-closed', cancel);
  window.addEventListener('online', () => {
    if (panelIsOpen()) schedule(900);
  });
  window.addEventListener('hera:adaptive-learning-sample', () => {
    if (panelIsOpen()) schedule(800);
  });

  window.HeraRecommendedTrafficWeather = {
    installed: true,
    version: VERSION,
    refresh: () => schedule(0),
    cancel,
    clearCache: () => {
      try {
        localStorage.removeItem(ROUTE_CACHE_KEY);
        localStorage.removeItem(WEATHER_CACHE_KEY);
      } catch (_) {}
    }
  };

  if (panelIsOpen()) schedule(300);
})();

const admin = require("firebase-admin");

const OPEN_METEO_URL = "https://api.open-meteo.com/v1/forecast";
const WEATHER_BATCH_SIZE = Number(process.env.WEATHER_BATCH_SIZE || 8);
const MAX_ACTIVE_IMPIANTI = Number(process.env.WEATHER_MAX_ACTIVE_IMPIANTI || 500);

exports.config = {
  schedule: "*/15 * * * *"
};

function getFirebaseCredential() {
  const raw = process.env.FIREBASE_SERVICE_ACCOUNT || process.env.GOOGLE_SERVICE_ACCOUNT_JSON || "";
  if (!raw.trim()) return admin.credential.applicationDefault();
  const parsed = JSON.parse(raw);
  if (parsed.private_key) parsed.private_key = parsed.private_key.replace(/\\n/g, "\n");
  return admin.credential.cert(parsed);
}

function getDatabaseURL() {
  return process.env.FIREBASE_DATABASE_URL || "https://hera-app-6cd2b-default-rtdb.firebaseio.com";
}

function ensureFirebaseApp() {
  if (admin.apps.length) return admin.app();
  return admin.initializeApp({
    credential: getFirebaseCredential(),
    databaseURL: getDatabaseURL()
  });
}

function toNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function getImpiantoCoordinates(impianto) {
  const lat = toNumber(impianto?.gpsY);
  const lon = toNumber(impianto?.gpsX);
  if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
  return { lat, lon };
}

function getPrecipitationAmount(values = {}) {
  return Math.max(toNumber(values.precipitation) || 0, toNumber(values.rain) || 0, toNumber(values.showers) || 0);
}

const RAIN_CODES = new Set([51, 53, 55, 56, 57, 61, 63, 65, 66, 67, 80, 81, 82]);
const THUNDER_CODES = new Set([95, 96, 99]);

function buildSlots(series = {}) {
  const times = Array.isArray(series.time) ? series.time : [];
  return times.map((time, index) => ({
    time,
    timestamp: new Date(time).getTime(),
    precipitation: series.precipitation?.[index],
    precipitation_probability: series.precipitation_probability?.[index],
    rain: series.rain?.[index],
    showers: series.showers?.[index],
    weather_code: series.weather_code?.[index],
    wind_speed_10m: series.wind_speed_10m?.[index],
    wind_gusts_10m: series.wind_gusts_10m?.[index]
  })).filter((slot) => Number.isFinite(slot.timestamp));
}

function getNextHourSlots(data = {}) {
  const now = Date.now();
  const nextHour = now + 60 * 60 * 1000;
  const minutely = buildSlots(data.minutely_15).filter((slot) => slot.timestamp >= now - 15 * 60 * 1000 && slot.timestamp <= nextHour);
  if (minutely.length) return minutely;
  return buildSlots(data.hourly).filter((slot) => slot.timestamp >= now - 60 * 60 * 1000 && slot.timestamp <= nextHour);
}

function summarizeWeather(data = {}) {
  const current = data.current || {};
  const currentCode = Number(current.weather_code);
  const currentPrecipitation = getPrecipitationAmount(current);
  const currentRain = currentPrecipitation > 0 || RAIN_CODES.has(currentCode);
  const slots = getNextHourSlots(data);
  const nextRain = slots.some((slot) => {
    const probability = Number(slot.precipitation_probability) || 0;
    return getPrecipitationAmount(slot) > 0 || probability >= 40 || RAIN_CODES.has(Number(slot.weather_code));
  });
  const thunder = THUNDER_CODES.has(currentCode) || slots.some((slot) => THUNDER_CODES.has(Number(slot.weather_code)));
  const wind = Number(current.wind_speed_10m) || 0;
  const gust = Number(current.wind_gusts_10m) || 0;
  const strongWind = wind >= 50 || gust >= 70 || slots.some((slot) => (Number(slot.wind_speed_10m) || 0) >= 50 || (Number(slot.wind_gusts_10m) || 0) >= 70);
  const severeRain = currentPrecipitation >= 5 || [65, 82, 96, 99].includes(currentCode) || slots.some((slot) => getPrecipitationAmount(slot) >= 5 || [65, 82, 96, 99].includes(Number(slot.weather_code)));

  const alertMessages = [];
  if (thunder) alertMessages.push("Temporale previsto");
  if (strongWind) alertMessages.push("Vento forte");
  if (severeRain) alertMessages.push("Precipitazioni intense");
  else if (currentRain || nextRain) alertMessages.push("Pioggia prevista");

  const status = thunder || strongWind || severeRain ? "red" : currentRain || nextRain ? "yellow" : "green";
  return {
    status,
    rain: currentRain || nextRain,
    alert: alertMessages.join("; "),
    temp: Number.isFinite(Number(current.temperature_2m)) ? Number(current.temperature_2m) : null,
    wind: Number.isFinite(wind) ? wind : null,
    updatedAt: Date.now()
  };
}

async function fetchWeatherForImpianto(impianto) {
  const coordinates = getImpiantoCoordinates(impianto);
  if (!coordinates) return null;
  const params = new URLSearchParams({
    latitude: String(coordinates.lat),
    longitude: String(coordinates.lon),
    current: "temperature_2m,precipitation,rain,showers,weather_code,wind_speed_10m,wind_gusts_10m",
    minutely_15: "precipitation,weather_code,wind_speed_10m,wind_gusts_10m",
    hourly: "precipitation_probability,precipitation,weather_code,wind_speed_10m,wind_gusts_10m",
    forecast_hours: "2",
    forecast_days: "1",
    timezone: "auto"
  });
  const response = await fetch(`${OPEN_METEO_URL}?${params.toString()}`, { headers: { Accept: "application/json" } });
  if (!response.ok) throw new Error(`Open-Meteo ${response.status}`);
  return summarizeWeather(await response.json());
}

async function loadActiveImpianti(db) {
  const commesseSnapshot = await db.collection("commesse").get();
  const impianti = [];
  for (const commessaDoc of commesseSnapshot.docs) {
    const snapshot = await commessaDoc.ref.collection("impianti").get();
    snapshot.docs.forEach((doc) => {
      if (impianti.length >= MAX_ACTIVE_IMPIANTI) return;
      const data = doc.data() || {};
      if (data.done === true) return;
      if (!getImpiantoCoordinates(data)) return;
      impianti.push({ id: doc.id, commessaId: commessaDoc.id, ...data });
    });
    if (impianti.length >= MAX_ACTIVE_IMPIANTI) break;
  }
  return impianti;
}

async function processBatch(items, callback, size = WEATHER_BATCH_SIZE) {
  const results = [];
  for (let index = 0; index < items.length; index += size) {
    const batch = items.slice(index, index + size);
    results.push(...await Promise.allSettled(batch.map(callback)));
  }
  return results;
}

exports.handler = async () => {
  ensureFirebaseApp();
  const db = admin.firestore();
  const rtdb = admin.database();
  const activeImpianti = await loadActiveImpianti(db);
  const updates = {};

  const results = await processBatch(activeImpianti, async (impianto) => {
    const weather = await fetchWeatherForImpianto(impianto);
    if (weather) updates[`weather/${impianto.id}`] = weather;
    return weather;
  });

  results.forEach((result, index) => {
    if (result.status === "rejected") {
      const impianto = activeImpianti[index];
      updates[`weather/${impianto.id}`] = {
        status: "unavailable",
        rain: false,
        alert: "Meteo non disponibile",
        temp: null,
        wind: null,
        updatedAt: Date.now()
      };
      console.warn("Meteo non disponibile", impianto.id, result.reason?.message || result.reason);
    }
  });

  if (Object.keys(updates).length) await rtdb.ref().update(updates);

  return {
    statusCode: 200,
    body: JSON.stringify({ ok: true, activeImpianti: activeImpianti.length, updated: Object.keys(updates).length })
  };
};

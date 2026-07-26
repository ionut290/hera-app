"use strict";

const OPEN_METEO_URL = "https://api.open-meteo.com/v1/forecast";
const WEATHER_TIMEOUT_MS = 10000;

function buildForecastParams(lat, lon, operational) {
  const params = new URLSearchParams({
    latitude: String(lat),
    longitude: String(lon),
    current: "temperature_2m,wind_speed_10m,weather_code",
    hourly: "temperature_2m,precipitation_probability,snowfall,visibility,weather_code,wind_speed_10m",
    forecast_days: "5",
    timezone: "auto"
  });

  if (operational) {
    params.set("current", "temperature_2m,apparent_temperature,relative_humidity_2m,precipitation,rain,showers,weather_code,wind_speed_10m,wind_direction_10m,wind_gusts_10m");
    params.set("minutely_15", "precipitation,weather_code");
    params.set("hourly", "temperature_2m,precipitation_probability,precipitation,rain,showers,snowfall,visibility,weather_code,apparent_temperature,wind_speed_10m,wind_direction_10m,wind_gusts_10m");
    params.set("forecast_hours", "12");
    params.set("forecast_minutely_15", "48");
    params.set("forecast_days", "1");
  }

  return params;
}

async function fetchForecast(lat, lon, operational) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), WEATHER_TIMEOUT_MS);
  const url = `${OPEN_METEO_URL}?${buildForecastParams(lat, lon, operational).toString()}`;
  try {
    const response = await fetch(url, {
      headers: { Accept: "application/json" },
      signal: controller.signal
    });
    if (!response.ok) throw new Error(`Open-Meteo HTTP ${response.status}`);
    const payload = await response.json();
    if (!payload?.current || !payload?.hourly) throw new Error("Risposta Open-Meteo non valida");
    return payload;
  } finally {
    clearTimeout(timeoutId);
  }
}

exports.handler = async (event) => {
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "public, max-age=120, s-maxage=300, stale-while-revalidate=900",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, OPTIONS"
  };

  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers, body: "" };
  if (event.httpMethod !== "GET") {
    return { statusCode: 405, headers, body: JSON.stringify({ ok: false, error: "Metodo non consentito" }) };
  }

  const lat = Number(event.queryStringParameters?.lat);
  const lon = Number(event.queryStringParameters?.lon);
  const operational = event.queryStringParameters?.operational === "1";
  if (!Number.isFinite(lat) || lat < -90 || lat > 90 || !Number.isFinite(lon) || lon < -180 || lon > 180) {
    return { statusCode: 400, headers, body: JSON.stringify({ ok: false, error: "Coordinate meteo non valide" }) };
  }

  try {
    const forecast = await fetchForecast(lat, lon, operational);
    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({ ...forecast, provider: "Open-Meteo proxy" })
    };
  } catch (error) {
    console.error("Proxy meteo non disponibile:", error);
    return {
      statusCode: 502,
      headers: { ...headers, "Cache-Control": "no-store" },
      body: JSON.stringify({ ok: false, error: "Servizio meteo temporaneamente non disponibile" })
    };
  }
};

exports.buildForecastParams = buildForecastParams;

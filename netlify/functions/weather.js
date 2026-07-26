"use strict";

const OPEN_METEO_PUBLIC_URL = "https://api.open-meteo.com/v1/forecast";
const OPEN_METEO_CUSTOMER_URL = "https://customer-api.open-meteo.com/v1/forecast";
const MET_NORWAY_URL = "https://api.met.no/weatherapi/locationforecast/2.0/complete";
const WEATHER_TIMEOUT_MS = 6500;
const MET_NORWAY_USER_AGENT = "VargaCantieri/1.0 github.com/ionut290/hera-app";

function buildForecastParams(lat, lon, operational, apiKey = "") {
  const params = new URLSearchParams({
    latitude: String(lat),
    longitude: String(lon),
    current: "temperature_2m,wind_speed_10m,weather_code",
    hourly: "temperature_2m,precipitation_probability,snowfall,visibility,weather_code,wind_speed_10m",
    forecast_days: "5",
    timezone: "auto",
    models: "best_match",
    cell_selection: "land",
    wind_speed_unit: "kmh"
  });

  if (operational) {
    params.set("current", "temperature_2m,apparent_temperature,relative_humidity_2m,precipitation,rain,showers,weather_code,wind_speed_10m,wind_direction_10m,wind_gusts_10m");
    params.set("minutely_15", "temperature_2m,relative_humidity_2m,apparent_temperature,precipitation,rain,showers,snowfall,weather_code,wind_speed_10m,wind_direction_10m,wind_gusts_10m,visibility,lightning_potential");
    params.set("hourly", "temperature_2m,relative_humidity_2m,dew_point_2m,apparent_temperature,precipitation_probability,precipitation,rain,showers,snowfall,visibility,weather_code,wind_speed_10m,wind_direction_10m,wind_gusts_10m,cape,uv_index");
    params.set("forecast_hours", "12");
    params.set("forecast_minutely_15", "48");
    params.set("forecast_days", "1");
  }

  if (apiKey) params.set("apikey", apiKey);
  return params;
}

function getOpenMeteoConfig() {
  const apiKey = String(process.env.OPEN_METEO_API_KEY || "").trim();
  return {
    apiKey,
    url: apiKey ? OPEN_METEO_CUSTOMER_URL : OPEN_METEO_PUBLIC_URL,
    provider: apiKey ? "Open-Meteo Best Match (customer)" : "Open-Meteo Best Match"
  };
}

async function fetchOpenMeteoForecast(lat, lon, operational) {
  const config = getOpenMeteoConfig();
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), WEATHER_TIMEOUT_MS);
  const url = `${config.url}?${buildForecastParams(lat, lon, operational, config.apiKey).toString()}`;
  try {
    const response = await fetch(url, {
      headers: { Accept: "application/json" },
      signal: controller.signal
    });
    if (!response.ok) throw new Error(`Open-Meteo HTTP ${response.status}`);
    const payload = await response.json();
    if (!payload?.current || !payload?.hourly) throw new Error("Risposta Open-Meteo non valida");
    return {
      ...payload,
      provider: config.provider,
      providerUrl: "https://open-meteo.com/",
      modelSelection: "best_match"
    };
  } finally {
    clearTimeout(timeoutId);
  }
}

function mapMetNorwaySymbolToWmo(symbolCode = "") {
  const symbol = String(symbolCode).toLowerCase();
  if (symbol.includes("thunder")) return symbol.includes("rain") || symbol.includes("sleet") || symbol.includes("snow") ? 96 : 95;
  if (symbol.includes("heavyrainshowers")) return 82;
  if (symbol.includes("rainshowers")) return 80;
  if (symbol.includes("heavyrain")) return 65;
  if (symbol.includes("lightrain")) return 61;
  if (symbol.includes("rain")) return 63;
  if (symbol.includes("heavysnowshowers")) return 86;
  if (symbol.includes("snowshowers")) return 85;
  if (symbol.includes("heavysnow")) return 75;
  if (symbol.includes("lightsnow")) return 71;
  if (symbol.includes("snow")) return 73;
  if (symbol.includes("sleet")) return 67;
  if (symbol.includes("fog")) return 45;
  if (symbol.includes("partlycloudy")) return 2;
  if (symbol.includes("cloudy")) return 3;
  if (symbol.includes("fair")) return 1;
  return 0;
}

function getMetNorwayPeriod(data = {}) {
  return data.next_1_hours || data.next_6_hours || data.next_12_hours || {};
}

function normalizeMetNorwayForecast(payload, operational) {
  const properties = payload?.properties || {};
  const units = properties.meta?.units || {};
  const sourceRows = Array.isArray(properties.timeseries) ? properties.timeseries : [];
  const rows = sourceRows.slice(0, operational ? 13 : 121);
  if (!rows.length) throw new Error("Risposta MET Norway non valida");

  const getInstant = (row) => row?.data?.instant?.details || {};
  const getPeriod = (row) => getMetNorwayPeriod(row?.data || {});
  const getDetails = (row) => getPeriod(row)?.details || {};
  const getSymbol = (row) => getPeriod(row)?.summary?.symbol_code || "";
  const getPrecipitation = (row) => Number(getDetails(row).precipitation_amount || 0);
  const firstInstant = getInstant(rows[0]);
  const firstDetails = getDetails(rows[0]);
  const firstCode = mapMetNorwaySymbolToWmo(getSymbol(rows[0]));

  const hourly = {
    time: rows.map((row) => row.time),
    temperature_2m: rows.map((row) => getInstant(row).air_temperature ?? null),
    relative_humidity_2m: rows.map((row) => getInstant(row).relative_humidity ?? null),
    precipitation_probability: rows.map((row) => getDetails(row).probability_of_precipitation ?? null),
    precipitation: rows.map(getPrecipitation),
    rain: rows.map((row) => /rain|sleet/i.test(getSymbol(row)) ? getPrecipitation(row) : 0),
    showers: rows.map((row) => /showers/i.test(getSymbol(row)) ? getPrecipitation(row) : 0),
    snowfall: rows.map((row) => /snow/i.test(getSymbol(row)) ? getPrecipitation(row) : 0),
    visibility: rows.map(() => null),
    weather_code: rows.map((row) => mapMetNorwaySymbolToWmo(getSymbol(row))),
    wind_speed_10m: rows.map((row) => {
      const speed = Number(getInstant(row).wind_speed);
      return Number.isFinite(speed) ? speed * 3.6 : null;
    }),
    wind_direction_10m: rows.map((row) => getInstant(row).wind_from_direction ?? null),
    wind_gusts_10m: rows.map((row) => {
      const gust = Number(getInstant(row).wind_speed_of_gust);
      return Number.isFinite(gust) ? gust * 3.6 : null;
    })
  };

  return {
    latitude: payload.geometry?.coordinates?.[1] ?? null,
    longitude: payload.geometry?.coordinates?.[0] ?? null,
    timezone: "UTC",
    current_units: {
      temperature_2m: units.air_temperature || "°C",
      relative_humidity_2m: units.relative_humidity || "%",
      precipitation: units.precipitation_amount || "mm",
      wind_speed_10m: "km/h",
      wind_direction_10m: units.wind_from_direction || "°",
      wind_gusts_10m: "km/h"
    },
    hourly_units: {
      temperature_2m: units.air_temperature || "°C",
      relative_humidity_2m: units.relative_humidity || "%",
      precipitation_probability: units.probability_of_precipitation || "%",
      precipitation: units.precipitation_amount || "mm",
      rain: units.precipitation_amount || "mm",
      showers: units.precipitation_amount || "mm",
      snowfall: units.precipitation_amount || "mm",
      visibility: "m",
      weather_code: "wmo code",
      wind_speed_10m: "km/h",
      wind_direction_10m: units.wind_from_direction || "°",
      wind_gusts_10m: "km/h"
    },
    current: {
      time: rows[0].time,
      temperature_2m: firstInstant.air_temperature ?? null,
      apparent_temperature: null,
      relative_humidity_2m: firstInstant.relative_humidity ?? null,
      precipitation: 0,
      rain: 0,
      showers: 0,
      weather_code: firstCode,
      wind_speed_10m: Number.isFinite(Number(firstInstant.wind_speed)) ? Number(firstInstant.wind_speed) * 3.6 : null,
      wind_direction_10m: firstInstant.wind_from_direction ?? null,
      wind_gusts_10m: Number.isFinite(Number(firstInstant.wind_speed_of_gust)) ? Number(firstInstant.wind_speed_of_gust) * 3.6 : null,
      precipitation_probability: firstDetails.probability_of_precipitation ?? null
    },
    hourly,
    provider: "MET Norway fallback",
    providerUrl: "https://api.met.no/weatherapi/locationforecast/2.0/documentation",
    modelSelection: "independent_fallback"
  };
}

async function fetchMetNorwayForecast(lat, lon, operational) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), WEATHER_TIMEOUT_MS);
  const params = new URLSearchParams({
    lat: Number(lat).toFixed(5),
    lon: Number(lon).toFixed(5)
  });
  try {
    const response = await fetch(`${MET_NORWAY_URL}?${params.toString()}`, {
      headers: {
        Accept: "application/json",
        "User-Agent": MET_NORWAY_USER_AGENT
      },
      signal: controller.signal
    });
    if (!response.ok) throw new Error(`MET Norway HTTP ${response.status}`);
    return normalizeMetNorwayForecast(await response.json(), operational);
  } finally {
    clearTimeout(timeoutId);
  }
}

async function fetchForecast(lat, lon, operational) {
  const providerErrors = [];
  try {
    return await fetchOpenMeteoForecast(lat, lon, operational);
  } catch (error) {
    providerErrors.push(`Open-Meteo: ${error?.message || error}`);
  }
  try {
    const fallback = await fetchMetNorwayForecast(lat, lon, operational);
    return { ...fallback, providerErrors };
  } catch (error) {
    providerErrors.push(`MET Norway: ${error?.message || error}`);
  }
  throw new Error(providerErrors.join("; "));
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
      body: JSON.stringify(forecast)
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

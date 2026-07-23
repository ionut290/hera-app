"use strict";

const core = require("../../fuel-stations-core.js");
const OVERPASS_ENDPOINTS = ["https://overpass-api.de/api/interpreter", "https://overpass.kumi.systems/api/interpreter"];

function haversine(lat1, lon1, lat2, lon2) {
  const toRad = (value) => value * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLon / 2) ** 2;
  return 6371 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

async function fetchJson(url, options, timeoutMs) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return response.json();
  } finally {
    clearTimeout(timeout);
  }
}

async function fromMimit(lat, lng, radius, fuel) {
  const payload = await fetchJson(core.MIMIT_API_URL, {
    method: "POST",
    headers: { Accept: "application/json", "Content-Type": "application/json" },
    body: JSON.stringify(core.buildMimitRequest(lat, lng, radius))
  }, 9000);
  if (!Array.isArray(payload?.results)) throw new Error("Risposta MIMIT non valida");
  return {
    stations: core.parseMimitStations(payload.results, fuel, { lat, lng }, haversine),
    source: "MIMIT Osservaprezzi",
    received: payload.results.length
  };
}

async function fromOverpass(lat, lng, radius, fuel) {
  const query = core.buildQuery(lat, lng, radius, fuel);
  let lastError;
  for (const endpoint of OVERPASS_ENDPOINTS) {
    try {
      const payload = await fetchJson(`${endpoint}?data=${encodeURIComponent(query)}`, { headers: { Accept: "application/json" } }, 12000);
      if (!Array.isArray(payload?.elements)) throw new Error("Risposta OpenStreetMap non valida");
      return {
        stations: core.parseStations(payload.elements, fuel, { lat, lng }, haversine)
          .map((station) => ({ ...station, source: "OpenStreetMap" })),
        source: "OpenStreetMap",
        received: payload.elements.length
      };
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("OpenStreetMap non disponibile");
}

exports.handler = async (event) => {
  const headers = {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "public, max-age=300, s-maxage=900, stale-while-revalidate=86400",
    "Access-Control-Allow-Origin": "*"
  };
  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers, body: "" };
  if (event.httpMethod !== "GET") return { statusCode: 405, headers, body: JSON.stringify({ ok: false, error: "Metodo non consentito" }) };
  const lat = Number(event.queryStringParameters?.lat);
  const lng = Number(event.queryStringParameters?.lng);
  const radius = Math.min(50, Math.max(1, Number(event.queryStringParameters?.radius) || 5));
  const fuel = String(event.queryStringParameters?.fuel || "");
  if (!Number.isFinite(lat) || !Number.isFinite(lng) || !core.FUEL_LABELS[fuel]) {
    return { statusCode: 400, headers, body: JSON.stringify({ ok: false, error: "Parametri ricerca non validi" }) };
  }
  const diagnostics = [];
  for (const source of [fromMimit, fromOverpass]) {
    try {
      const result = await source(lat, lng, radius, fuel);
      const stations = result.stations.filter((station) => station.distance <= radius).sort((a, b) => a.distance - b.distance);
      if (stations.length || source === fromOverpass) {
        return { statusCode: 200, headers, body: JSON.stringify({ ok: true, ...result, stations, diagnostics }) };
      }
      diagnostics.push(`${result.source}: nessun risultato compatibile`);
    } catch (error) {
      diagnostics.push(error?.message || String(error));
    }
  }
  return { statusCode: 503, headers, body: JSON.stringify({ ok: false, error: "Fonti distributori non disponibili", diagnostics }) };
};

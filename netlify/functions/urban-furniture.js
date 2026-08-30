"use strict";

const OVERPASS_ENDPOINTS = [
  "https://overpass-api.de/api/interpreter",
  "https://overpass.kumi.systems/api/interpreter"
];
const REQUEST_TIMEOUT_MS = 18000;
const MAX_RESULTS = 500;
const CATEGORIES = Object.freeze({
  bench: '["amenity"="bench"]',
  waste_basket: '["amenity"="waste_basket"]',
  drinking_water: '["amenity"="drinking_water"]',
  fountain: '["amenity"="fountain"]',
  picnic_table: '["leisure"="picnic_table"]',
  bicycle_parking: '["amenity"="bicycle_parking"]',
  toilets: '["amenity"="toilets"]',
  shelter: '["amenity"="shelter"]',
  street_lamp: '["highway"="street_lamp"]',
  playground: '["leisure"="playground"]',
  dog_park: '["leisure"="dog_park"]',
  fitness_station: '["leisure"="fitness_station"]',
  outdoor_seating: '["leisure"="outdoor_seating"]',
  bbq: '["amenity"="bbq"]',
  dog_toilet: '["amenity"="dog_toilet"]',
  recycling: '["amenity"="recycling"]',
  waste_disposal: '["amenity"="waste_disposal"]',
  post_box: '["amenity"="post_box"]',
  parcel_locker: '["amenity"="parcel_locker"]',
  telephone: '["amenity"="telephone"]',
  public_bookcase: '["amenity"="public_bookcase"]',
  bicycle_repair_station: '["amenity"="bicycle_repair_station"]',
  charging_station: '["amenity"="charging_station"]',
  motorcycle_parking: '["amenity"="motorcycle_parking"]',
  taxi: '["amenity"="taxi"]',
  compressed_air: '["amenity"="compressed_air"]',
  shower: '["amenity"="shower"]',
  water_point: '["amenity"="water_point"]',
  clock: '["amenity"="clock"]',
  grit_bin: '["amenity"="grit_bin"]',
  lounger: '["amenity"="lounger"]',
  give_box: '["amenity"="give_box"]',
  fire_hydrant: '["emergency"="fire_hydrant"]',
  defibrillator: '["emergency"="defibrillator"]',
  bollard: '["barrier"="bollard"]',
  cycle_barrier: '["barrier"="cycle_barrier"]',
  bus_stop: '["highway"="bus_stop"]',
  artwork: '["tourism"="artwork"]',
  information: '["tourism"="information"]'
});

function response(statusCode, body, cache = false) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": cache ? "public, max-age=120, s-maxage=300, stale-while-revalidate=900" : "no-store",
      "Access-Control-Allow-Origin": "*"
    },
    body: JSON.stringify(body)
  };
}

function escapeString(value) {
  return String(value || "").replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function clauses(scope, category) {
  const selected = category === "all" ? Object.values(CATEGORIES) : [CATEGORIES[category]];
  return selected.map((clause) => `nwr${clause}(${scope});`).join("");
}

function buildQuery(params) {
  const mode = String(params.mode || "municipality");
  const category = String(params.category || "all");
  if (category !== "all" && !CATEGORIES[category]) throw new Error("Tipo di arredo urbano non valido.");
  if (mode === "municipality") {
    const municipality = String(params.municipality || "").trim();
    if (municipality.length < 2 || municipality.length > 80) throw new Error("Nome del Comune non valido.");
    const town = escapeString(municipality);
    return `[out:json][timeout:25];area["boundary"="administrative"]["admin_level"="8"]["name"="${town}"]->.municipality;(${clauses("area.municipality", category)});out center tags ${MAX_RESULTS};`;
  }
  if (mode === "nearby") {
    const lat = Number(params.lat);
    const lon = Number(params.lon);
    if (!Number.isFinite(lat) || !Number.isFinite(lon) || lat < -90 || lat > 90 || lon < -180 || lon > 180) throw new Error("Coordinate posizione non valide.");
    return `[out:json][timeout:25];(${clauses(`around:3000,${lat.toFixed(6)},${lon.toFixed(6)}`, category)});out center tags ${MAX_RESULTS};`;
  }
  throw new Error("Modalità di ricerca non valida.");
}

async function fetchEndpoint(endpoint, query) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    const response = await fetch(`${endpoint}?data=${encodeURIComponent(query)}`, {
      headers: { Accept: "application/json", "User-Agent": "VargaCantieri/1.0" },
      signal: controller.signal
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const payload = await response.json();
    if (!Array.isArray(payload?.elements)) throw new Error("Risposta OpenStreetMap non valida");
    return payload;
  } finally {
    clearTimeout(timeout);
  }
}

async function fetchOverpass(query) {
  try {
    return await Promise.any(OVERPASS_ENDPOINTS.map((endpoint) => fetchEndpoint(endpoint, query)));
  } catch (error) {
    const reasons = error?.errors?.map((item) => item?.message || String(item)).filter(Boolean) || [];
    throw new Error(reasons.length ? reasons.join("; ") : "Server OpenStreetMap non raggiungibili");
  }
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers: { "Access-Control-Allow-Origin": "*" }, body: "" };
  if (event.httpMethod !== "GET") return response(405, { ok: false, error: "Metodo non consentito." });
  let query;
  try {
    query = buildQuery(event.queryStringParameters || {});
  } catch (error) {
    return response(400, { ok: false, error: error.message || "Parametri non validi." });
  }
  try {
    const payload = await fetchOverpass(query);
    return response(200, { ok: true, elements: payload.elements.slice(0, MAX_RESULTS) }, true);
  } catch (error) {
    console.warn("Urban furniture Overpass unavailable:", error?.message || error);
    return response(503, { ok: false, error: "Il catasto OpenStreetMap è temporaneamente occupato. Riprova tra qualche secondo." });
  }
};

exports._test = { buildQuery, CATEGORIES };

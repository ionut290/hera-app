"use strict";

const OVERPASS_ENDPOINTS = ["https://overpass-api.de/api/interpreter", "https://overpass.kumi.systems/api/interpreter"];
const REQUEST_TIMEOUT_MS = 22000;
const MAX_RESULTS = 1000;

function response(statusCode, body, cache = false) {
  return { statusCode, headers: { "Content-Type": "application/json; charset=utf-8", "Cache-Control": cache ? "public, max-age=300, s-maxage=1800, stale-while-revalidate=3600" : "no-store", "Access-Control-Allow-Origin": "*" }, body: JSON.stringify(body) };
}

function escapeString(value) {
  return String(value || "").replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}

function buildQuery(params) {
  const area = String(params.area || "Bologna").trim();
  const level = area.toLocaleLowerCase("it-IT") === "emilia-romagna" ? "4" : "6";
  if (area.length < 2 || area.length > 80) throw new Error("Territorio non valido.");
  const safeArea = escapeString(area);
  return `[out:json][timeout:30];area["boundary"="administrative"]["admin_level"="${level}"]["name"="${safeArea}"]->.searchArea;nwr["man_made"="pumping_station"](area.searchArea);out center tags ${MAX_RESULTS};`;
}

function isRelevantStation(element) {
  const tags = element?.tags || {};
  const declared = `${tags.substance || ""} ${tags.pumping_station || ""}`.toLocaleLowerCase("it-IT");
  const name = String(tags.name || "").toLocaleLowerCase("it-IT");
  if (/(sewage|wastewater|fogn|reflu|sollev)/.test(`${declared} ${name}`)) return true;
  if (/(^|\s)(lpg|gas|oil|petroleum|fuel|drinking_water|water)(\s|$)/.test(declared)) return false;
  return !/(acqua|acquedott|bonifica|chiavica|irrig)/.test(name);
}

async function fetchEndpoint(endpoint, query) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    const result = await fetch(`${endpoint}?data=${encodeURIComponent(query)}`, { headers: { Accept: "application/json", "User-Agent": "VargaCantieri/1.0" }, signal: controller.signal });
    if (!result.ok) throw new Error(`HTTP ${result.status}`);
    const payload = await result.json();
    if (!Array.isArray(payload?.elements)) throw new Error("Risposta OpenStreetMap non valida");
    return payload.elements;
  } finally { clearTimeout(timeout); }
}

async function fetchOverpass(query) {
  try { return await Promise.any(OVERPASS_ENDPOINTS.map((endpoint) => fetchEndpoint(endpoint, query))); }
  catch (error) {
    const reasons = error?.errors?.map((item) => item?.message || String(item)).filter(Boolean) || [];
    throw new Error(reasons.length ? reasons.join("; ") : "Server OpenStreetMap non raggiungibili");
  }
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers: { "Access-Control-Allow-Origin": "*" }, body: "" };
  if (event.httpMethod !== "GET") return response(405, { ok: false, error: "Metodo non consentito." });
  let query;
  try { query = buildQuery(event.queryStringParameters || {}); }
  catch (error) { return response(400, { ok: false, error: error.message || "Parametri non validi." }); }
  try {
    const elements = (await fetchOverpass(query)).filter(isRelevantStation);
    return response(200, { ok: true, elements: elements.slice(0, MAX_RESULTS) }, true);
  } catch (error) {
    console.warn("Wastewater lift stations Overpass unavailable:", error?.message || error);
    return response(503, { ok: false, error: "OpenStreetMap è temporaneamente occupato. Riprova tra qualche secondo." });
  }
};

exports._test = { buildQuery, isRelevantStation };

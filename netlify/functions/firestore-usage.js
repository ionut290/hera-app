"use strict";

const crypto = require("crypto");
const { authenticateEvent } = require("./_firebase-token");

const PROJECT_ID = String(process.env.FIREBASE_PROJECT_ID || "hera-app-6cd2b").trim();
const CACHE_MS = 15 * 60 * 1000;
const METRICS = {
  reads: "firestore.googleapis.com/document/read_ops_count",
  writes: "firestore.googleapis.com/document/write_ops_count",
  deletes: "firestore.googleapis.com/document/delete_ops_count"
};
const LIMITS = { reads: 50000, writes: 20000, deletes: 20000 };
let cache = { expiresAt: 0, value: null };
let tokenCache = { expiresAt: 0, accessToken: "" };

function json(statusCode, payload) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, OPTIONS",
      "Access-Control-Allow-Headers": "Authorization, Content-Type"
    },
    body: JSON.stringify(payload)
  };
}

function base64Url(value) {
  return Buffer.from(value).toString("base64").replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_");
}

function serviceAccount() {
  const raw = process.env.GOOGLE_CLOUD_SERVICE_ACCOUNT_JSON
    || process.env.FIREBASE_SERVICE_ACCOUNT_JSON
    || process.env.GOOGLE_SERVICE_ACCOUNT_JSON
    || "";
  if (!raw) throw new Error("Credenziali Cloud Monitoring non configurate su Netlify.");
  let parsed;
  try { parsed = JSON.parse(raw); } catch (_) { throw new Error("JSON del service account non valido."); }
  if (!parsed.client_email || !parsed.private_key) throw new Error("Service account incompleto.");
  return parsed;
}

async function accessToken() {
  if (tokenCache.accessToken && Date.now() < tokenCache.expiresAt) return tokenCache.accessToken;
  const account = serviceAccount();
  const now = Math.floor(Date.now() / 1000);
  const header = base64Url(JSON.stringify({ alg: "RS256", typ: "JWT" }));
  const claims = base64Url(JSON.stringify({
    iss: account.client_email,
    scope: "https://www.googleapis.com/auth/monitoring.read",
    aud: "https://oauth2.googleapis.com/token",
    iat: now,
    exp: now + 3600
  }));
  const unsigned = `${header}.${claims}`;
  const signature = crypto.sign("RSA-SHA256", Buffer.from(unsigned), account.private_key);
  const assertion = `${unsigned}.${base64Url(signature)}`;
  const response = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion })
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok || !data.access_token) throw new Error(data.error_description || "Autorizzazione Cloud Monitoring non riuscita.");
  tokenCache = { accessToken: data.access_token, expiresAt: Date.now() + Math.max(300, Number(data.expires_in || 3600) - 120) * 1000 };
  return tokenCache.accessToken;
}

function pacificMidnightUtc(now = new Date()) {
  const zone = "America/Los_Angeles";
  const parts = Object.fromEntries(new Intl.DateTimeFormat("en-US", {
    timeZone: zone, year: "numeric", month: "2-digit", day: "2-digit"
  }).formatToParts(now).filter((part) => part.type !== "literal").map((part) => [part.type, Number(part.value)]));
  let guess = Date.UTC(parts.year, parts.month - 1, parts.day, 0, 0, 0);
  for (let i = 0; i < 2; i += 1) {
    const zoned = Object.fromEntries(new Intl.DateTimeFormat("en-US", {
      timeZone: zone, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23"
    }).formatToParts(new Date(guess)).filter((part) => part.type !== "literal").map((part) => [part.type, Number(part.value)]));
    const represented = Date.UTC(zoned.year, zoned.month - 1, zoned.day, zoned.hour, zoned.minute, zoned.second);
    guess += Date.UTC(parts.year, parts.month - 1, parts.day, 0, 0, 0) - represented;
  }
  return new Date(guess);
}

function pointValue(point) {
  const value = point?.value || {};
  return Number(value.int64Value ?? value.doubleValue ?? 0) || 0;
}

function pointBelongsToPeriod(point, start, end) {
  const pointEnd = new Date(point?.interval?.endTime || 0).getTime();
  return Number.isFinite(pointEnd) && pointEnd > start.getTime() && pointEnd <= end.getTime();
}

async function readMetric(metricType, start, end, token) {
  let total = 0;
  let pageToken = "";

  do {
    const params = new URLSearchParams({
      filter: `metric.type="${metricType}"`,
      "interval.startTime": start.toISOString(),
      "interval.endTime": end.toISOString(),
      view: "FULL",
      "aggregation.alignmentPeriod": "60s",
      "aggregation.perSeriesAligner": "ALIGN_SUM",
      "aggregation.crossSeriesReducer": "REDUCE_SUM",
      pageSize: "1000"
    });
    if (pageToken) params.set("pageToken", pageToken);

    const response = await fetch(`https://monitoring.googleapis.com/v3/projects/${encodeURIComponent(PROJECT_ID)}/timeSeries?${params}`, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json" }
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data?.error?.message || `Cloud Monitoring HTTP ${response.status}`);

    total += (data.timeSeries || []).reduce((sum, series) => sum + (series.points || [])
      .filter((point) => pointBelongsToPeriod(point, start, end))
      .reduce((subtotal, point) => subtotal + pointValue(point), 0), 0);
    pageToken = String(data.nextPageToken || "");
  } while (pageToken);

  return total;
}

function allowedAdmin(user) {
  const allowed = new Set(String(process.env.FIRESTORE_USAGE_ALLOWED_EMAILS || "ionut29019@gmail.com")
    .split(/[;,\s]+/).map((email) => email.trim().toLowerCase()).filter(Boolean));
  return Boolean(user?.email && allowed.has(String(user.email).toLowerCase()));
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return json(204, {});
  if (event.httpMethod !== "GET") return json(405, { ok: false, error: "Metodo non consentito." });
  try {
    const user = await authenticateEvent(event);
    if (!allowedAdmin(user)) return json(403, { ok: false, error: "Funzione disponibile solo all'amministratore." });
    if (cache.value && Date.now() < cache.expiresAt) return json(200, { ...cache.value, cached: true });

    const end = new Date();
    const start = pacificMidnightUtc(end);
    const token = await accessToken();
    const [reads, writes, deletes] = await Promise.all([
      readMetric(METRICS.reads, start, end, token),
      readMetric(METRICS.writes, start, end, token),
      readMetric(METRICS.deletes, start, end, token)
    ]);
    const usage = { reads: Math.round(reads), writes: Math.round(writes), deletes: Math.round(deletes) };
    const percentages = Object.fromEntries(Object.keys(usage).map((key) => [key, Math.min(999, (usage[key] / LIMITS[key]) * 100)]));
    const value = {
      ok: true,
      projectId: PROJECT_ID,
      usage,
      limits: LIMITS,
      percentages,
      periodStart: start.toISOString(),
      periodEnd: end.toISOString(),
      generatedAt: new Date().toISOString(),
      source: "Google Cloud Monitoring — intervallo esatto del giorno Firestore",
      cached: false
    };
    cache = { value, expiresAt: Date.now() + CACHE_MS };
    return json(200, value);
  } catch (error) {
    console.error("firestore-usage:", error);
    return json(500, { ok: false, error: error?.message || "Consumo Firestore non disponibile." });
  }
};

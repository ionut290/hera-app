"use strict";

const crypto = require("crypto");
const { authenticateEvent } = require("./_firebase-token");

const PROJECT_ID = String(process.env.FIREBASE_PROJECT_ID || "hera-app-6cd2b").trim();
const CACHE_MS = 10 * 60 * 1000;
const METRIC = "serviceruntime.googleapis.com/api/request_count";
let cache = { expiresAt: 0, value: null };
let tokenCache = { expiresAt: 0, accessToken: "" };

const FRIENDLY_NAMES = {
  "maps-backend.googleapis.com": "Maps JavaScript API",
  "places.googleapis.com": "Places API (New)",
  "street-view-image-backend.googleapis.com": "Street View Static API",
  "streetviewpublish.googleapis.com": "Street View Publish API",
  "maps-android-backend.googleapis.com": "Maps SDK for Android",
  "maps-ios-backend.googleapis.com": "Maps SDK for iOS",
  "identitytoolkit.googleapis.com": "Firebase Authentication",
  "firebaseinstallations.googleapis.com": "Firebase Installations",
  "firebaseappcheck.googleapis.com": "Firebase App Check",
  "firestore.googleapis.com": "Cloud Firestore API",
  "firebase.googleapis.com": "Firebase Management API",
  "fcm.googleapis.com": "Firebase Cloud Messaging",
  "cloudfunctions.googleapis.com": "Cloud Functions API",
  "storage.googleapis.com": "Cloud Storage API"
};

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

function allowedAdmin(user) {
  const allowed = new Set(String(process.env.FIRESTORE_USAGE_ALLOWED_EMAILS || "ionut29019@gmail.com")
    .split(/[;,\s]+/).map((email) => email.trim().toLowerCase()).filter(Boolean));
  return Boolean(user?.email && allowed.has(String(user.email).toLowerCase()));
}

function monthStartUtc(now = new Date()) {
  return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1, 0, 0, 0));
}

function pointValue(point) {
  const value = point?.value || {};
  return Number(value.int64Value ?? value.doubleValue ?? 0) || 0;
}

function friendlyName(service) {
  if (FRIENDLY_NAMES[service]) return FRIENDLY_NAMES[service];
  return String(service || "API sconosciuta")
    .replace(/\.googleapis\.com$/i, "")
    .replace(/[-_.]+/g, " ")
    .replace(/\b\w/g, (m) => m.toUpperCase());
}

async function readApiUsage(start, end, token) {
  const totals = new Map();
  let pageToken = "";
  do {
    const params = new URLSearchParams({
      filter: `metric.type="${METRIC}" AND resource.type="consumed_api"`,
      "interval.startTime": start.toISOString(),
      "interval.endTime": end.toISOString(),
      view: "FULL",
      "aggregation.alignmentPeriod": "3600s",
      "aggregation.perSeriesAligner": "ALIGN_SUM",
      pageSize: "1000"
    });
    if (pageToken) params.set("pageToken", pageToken);
    const response = await fetch(`https://monitoring.googleapis.com/v3/projects/${encodeURIComponent(PROJECT_ID)}/timeSeries?${params}`, {
      headers: { Authorization: `Bearer ${token}`, Accept: "application/json" }
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data?.error?.message || `Cloud Monitoring HTTP ${response.status}`);
    for (const series of data.timeSeries || []) {
      const service = String(series?.resource?.labels?.service || series?.metric?.labels?.service || "unknown");
      const amount = (series.points || []).reduce((sum, point) => sum + pointValue(point), 0);
      totals.set(service, (totals.get(service) || 0) + amount);
    }
    pageToken = String(data.nextPageToken || "");
  } while (pageToken);

  return [...totals.entries()]
    .map(([service, requests]) => ({ service, name: friendlyName(service), requests: Math.round(requests) }))
    .filter((item) => item.requests > 0)
    .sort((a, b) => b.requests - a.requests);
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return json(204, {});
  if (event.httpMethod !== "GET") return json(405, { ok: false, error: "Metodo non consentito." });
  try {
    const user = await authenticateEvent(event);
    if (!allowedAdmin(user)) return json(403, { ok: false, error: "Funzione disponibile solo all'amministratore." });
    if (cache.value && Date.now() < cache.expiresAt) return json(200, { ...cache.value, cached: true });

    const end = new Date();
    const start = monthStartUtc(end);
    const token = await accessToken();
    const services = await readApiUsage(start, end, token);
    const totalRequests = services.reduce((sum, item) => sum + item.requests, 0);
    const value = {
      ok: true,
      projectId: PROJECT_ID,
      period: "month",
      periodStart: start.toISOString(),
      periodEnd: end.toISOString(),
      totalRequests,
      services,
      externalProviders: [
        { name: "Open-Meteo", tracking: "proxy Netlify", note: "Non incluso nei contatori Google Cloud" },
        { name: "MET Norway", tracking: "fallback meteo", note: "Non incluso nei contatori Google Cloud" }
      ],
      generatedAt: new Date().toISOString(),
      source: "Google Cloud Monitoring — Service Runtime API request count",
      cached: false
    };
    cache = { value, expiresAt: Date.now() + CACHE_MS };
    return json(200, value);
  } catch (error) {
    console.error("api-usage:", error);
    return json(500, { ok: false, error: error?.message || "Consumo API non disponibile." });
  }
};

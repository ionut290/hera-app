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

// Riferimenti Google Maps Platform global pricing, aggiornati ad agosto 2026.
// Queste soglie servono per AVVISI nel Centro di Controllo: Service Runtime request_count
// non coincide necessariamente con gli eventi fatturabili di uno SKU, quindi non viene
// usato come hard-stop. Street View 360° mantiene invece il contatore applicativo esatto
// e il blocco separato a 4.800/mese.
const COST_POLICIES = {
  "maps-backend.googleapis.com": {
    sku: "Dynamic Maps",
    freeCap: 10000,
    warningAt: 8000,
    dangerAt: 9500,
    pricePer1000Usd: 7,
    hardBlock: false,
    note: "Cap gratuito Dynamic Maps: 10.000 eventi/mese. Conteggio Service Runtime indicativo."
  },
  "places.googleapis.com": {
    sku: "Places (prudenziale: Place Details Pro)",
    freeCap: 5000,
    warningAt: 4000,
    dangerAt: 4800,
    pricePer1000Usd: 17,
    hardBlock: false,
    note: "L'app richiede displayName, che può attivare Place Details Pro: usata soglia prudenziale 5.000/mese. Autocomplete Requests ha cap gratuito 10.000/mese."
  },
  "street-view-image-backend.googleapis.com": {
    sku: "Static Street View",
    freeCap: 10000,
    warningAt: 8000,
    dangerAt: 9500,
    pricePer1000Usd: 7,
    hardBlock: false,
    note: "Cap gratuito Static Street View: 10.000 eventi/mese."
  },
  "maps-android-backend.googleapis.com": {
    sku: "Maps SDK",
    freeCap: null,
    unlimitedFree: true,
    hardBlock: false,
    note: "Maps SDK indicato da Google con free usage cap illimitato."
  },
  "maps-ios-backend.googleapis.com": {
    sku: "Maps SDK",
    freeCap: null,
    unlimitedFree: true,
    hardBlock: false,
    note: "Maps SDK indicato da Google con free usage cap illimitato."
  }
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
    body: new URLSearchParams({ grant_type: "urn:ietf:params:oauth2:grant-type:jwt-bearer", assertion })
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

function policyStatus(requests, policy) {
  if (!policy) return { tone: "neutral", label: "ℹ️ Monitoraggio", percent: null, remaining: null };
  if (policy.unlimitedFree) return { tone: "safe", label: "🟢 Cap gratuito illimitato", percent: 0, remaining: null };
  const cap = Number(policy.freeCap || 0);
  if (!cap) return { tone: "neutral", label: "ℹ️ Nessuna soglia automatica", percent: null, remaining: null };
  const percent = Math.min(999, requests / cap * 100);
  const remaining = Math.max(0, cap - requests);
  if (requests >= cap) return { tone: "danger", label: "🔴 Oltre riferimento gratuito", percent, remaining };
  if (requests >= Number(policy.dangerAt || cap * 0.95)) return { tone: "danger", label: "🔴 Quasi al limite", percent, remaining };
  if (requests >= Number(policy.warningAt || cap * 0.8)) return { tone: "warning", label: "🟡 Attenzione", percent, remaining };
  return { tone: "safe", label: "🟢 Consumo regolare", percent, remaining };
}

function applyCostPolicy(item) {
  const policy = COST_POLICIES[item.service] || null;
  const status = policyStatus(item.requests, policy);
  const capLabel = policy?.unlimitedFree
    ? "illimitato"
    : policy?.freeCap
      ? new Intl.NumberFormat("it-IT").format(policy.freeCap)
      : "n/d";
  return {
    ...item,
    baseName: item.name,
    name: `${item.name} · ${status.label}${policy ? ` · riferimento ${capLabel}/mese` : ""}`,
    costPolicy: policy ? { ...policy, ...status } : { ...status, hardBlock: false },
    pricingGuard: policy?.hardBlock ? "hard" : "monitor-only"
  };
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
    .map(applyCostPolicy)
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
      costPolicyUpdatedAt: "2026-08-23",
      costPolicyMode: "monitor-only-except-street-view-360",
      policyNotice: "Le soglie API Google sono avvisi prudenti basati su Service Runtime e non bloccano funzioni critiche. Street View 360° resta l'unico hard-stop a 4.800/mese perché usa un contatore applicativo condiviso esatto.",
      externalProviders: [
        { name: "Open-Meteo", tracking: "proxy Netlify", note: "Non incluso nei contatori Google Cloud; nessun blocco automatico" },
        { name: "MET Norway", tracking: "fallback meteo", note: "Non incluso nei contatori Google Cloud; nessun blocco automatico" }
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

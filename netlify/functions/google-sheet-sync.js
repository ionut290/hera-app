"use strict";

const { authenticateEvent } = require("./_firebase-token");

const MAX_BODY_BYTES = 8 * 1024 * 1024;
const MAX_ROWS = 20000;
const MAX_COLUMNS = 80;
const REQUEST_TIMEOUT_MS = 25000;

function json(statusCode, payload, extraHeaders = {}) {
  return {
    statusCode,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization",
      ...extraHeaders
    },
    body: JSON.stringify(payload)
  };
}

function parseBody(event) {
  const raw = event.isBase64Encoded
    ? Buffer.from(event.body || "", "base64").toString("utf8")
    : String(event.body || "");
  if (!raw || Buffer.byteLength(raw, "utf8") > MAX_BODY_BYTES) {
    throw new Error("Richiesta vuota o troppo grande (massimo 8 MB).");
  }
  try {
    return JSON.parse(raw);
  } catch (_) {
    throw new Error("Dati di sincronizzazione non validi.");
  }
}

function validateGoogleSheetUrl(rawUrl) {
  const parsed = new URL(String(rawUrl || "").trim());
  if (parsed.protocol !== "https:" || parsed.hostname !== "docs.google.com") {
    throw new Error("Link Google Sheet non valido.");
  }
  const match = parsed.pathname.match(/^\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  if (!match) throw new Error("ID del Google Sheet non riconosciuto.");
  return match[1];
}

function validateAppsScriptUrl(rawUrl) {
  const parsed = new URL(String(rawUrl || "").trim());
  const allowedHost = parsed.hostname === "script.google.com" || parsed.hostname.endsWith(".script.google.com");
  if (parsed.protocol !== "https:" || !allowedHost || !/\/macros\/s\//.test(parsed.pathname)) {
    throw new Error("GOOGLE_SHEET_APPS_SCRIPT_URL non valido.");
  }
  return parsed.href;
}

function allowedEmails() {
  return new Set(String(process.env.GOOGLE_SHEET_SYNC_ALLOWED_EMAILS || "")
    .split(/[;,\s]+/)
    .map((email) => email.trim().toLowerCase())
    .filter(Boolean));
}

function validatePayload(payload) {
  const action = String(payload?.action || "");
  if (action === "createRegistrySpreadsheet") {
    const registry = ["personale", "mezzi"].includes(String(payload.registry)) ? String(payload.registry) : "";
    const sheetNames = Array.isArray(payload.sheetNames) ? payload.sheetNames.map(String) : [];
    if (!registry || !["PERSONALE", "MEZZI", "COMMESSE_PERSONALE", "COMMESSE_MEZZI", "LOG_SINCRONIZZAZIONE"].every((name) => sheetNames.includes(name))) throw new Error("Configurazione matrice Google non valida.");
    const sheetUrl = String(payload.sheetUrl || "").trim();
    return { action, registry, sheetNames, sheetUrl, spreadsheetId: sheetUrl ? validateGoogleSheetUrl(sheetUrl) : "" };
  }
  if (action === "syncRegistrySpreadsheet") {
    const spreadsheetId = validateGoogleSheetUrl(payload.sheetUrl);
    if (String(payload.spreadsheetId || spreadsheetId) !== spreadsheetId) throw new Error("Il link e l'ID del foglio non coincidono.");
    const sheets = payload.sheets && typeof payload.sheets === "object" ? payload.sheets : {};
    const required = ["PERSONALE", "MEZZI", "COMMESSE_PERSONALE", "COMMESSE_MEZZI", "LOG_SINCRONIZZAZIONE"];
    if (!required.every((name) => Array.isArray(sheets[name])) || Object.values(sheets).some((rows) => rows.length > MAX_ROWS)) throw new Error("Fogli registro non validi.");
    return {
      action,
      registry: String(payload.registry || ""),
      spreadsheetId,
      sheetUrl: payload.sheetUrl,
      gid: /^\d+$/.test(String(payload.gid || "0")) ? String(payload.gid || "0") : "0",
      sheets,
      conflictPolicy: payload.conflictPolicy === "APP_WINS" ? "APP_WINS" : "LATEST_WINS",
      noAutomaticDeletion: true
    };
  }
  if (action === "createSpreadsheet") {
    const headers = Array.isArray(payload.headers) ? payload.headers.map((value) => String(value ?? "")) : [];
    if (!String(payload.commessaId || "").trim()) throw new Error("ID commessa mancante.");
    if (!headers.includes("SYNC_KEY") || !headers.includes("IMPIANTO_KEY") || headers.length > MAX_COLUMNS) {
      throw new Error("Intestazioni del foglio non valide.");
    }
    return { action, commessaId: String(payload.commessaId).slice(0, 200), commessaName: String(payload.commessaName || "Senza nome").slice(0, 300), headers };
  }
  if (action !== "replaceRows") throw new Error("Azione di sincronizzazione non supportata.");
  const spreadsheetIdFromUrl = validateGoogleSheetUrl(payload.sheetUrl);
  const spreadsheetId = String(payload.spreadsheetId || spreadsheetIdFromUrl).trim();
  if (spreadsheetId !== spreadsheetIdFromUrl) throw new Error("Il link e l'ID del foglio non coincidono.");
  const gid = /^\d+$/.test(String(payload.gid || "0")) ? String(payload.gid || "0") : "0";
  const headers = Array.isArray(payload.headers) ? payload.headers.map((value) => String(value ?? "")) : [];
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!headers.length || headers.length > MAX_COLUMNS) throw new Error("Intestazioni del foglio non valide.");
  if (rows.length > MAX_ROWS) throw new Error(`Troppe righe: massimo ${MAX_ROWS}.`);
  if (rows.some((row) => !Array.isArray(row) || row.length > headers.length)) {
    throw new Error("Una o più righe del foglio non sono valide.");
  }
  return {
    action: "replaceRows",
    spreadsheetId,
    gid,
    headers,
    rows: rows.map((row) => headers.map((_, index) => row[index] ?? "")),
    commessaId: String(payload.commessaId || "").slice(0, 200),
    commessaName: String(payload.commessaName || "").slice(0, 300),
    operationId: String(payload.operationId || "").slice(0, 300)
  };
}

async function callAppsScript(payload, user) {
  const scriptUrl = validateAppsScriptUrl(process.env.GOOGLE_SHEET_APPS_SCRIPT_URL);
  const secret = String(process.env.GOOGLE_SHEET_SYNC_SECRET || "").trim();
  if (!secret) throw new Error("GOOGLE_SHEET_SYNC_SECRET non configurato.");
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
  try {
    const response = await fetch(scriptUrl, {
      method: "POST",
      redirect: "follow",
      signal: controller.signal,
      headers: { "Content-Type": "text/plain;charset=utf-8", Accept: "application/json" },
      body: JSON.stringify({
        ...payload,
        secret,
        requestedByUid: user.sub,
        requestedByEmail: user.email || "",
        requestedAt: new Date().toISOString()
      })
    });
    const text = await response.text();
    let result;
    try { result = JSON.parse(text); } catch (_) { result = null; }
    if (!response.ok || !result?.ok) {
      throw new Error(result?.error || `Apps Script HTTP ${response.status}`);
    }
    return result;
  } catch (error) {
    if (error?.name === "AbortError") throw new Error("Apps Script non ha risposto in tempo.");
    throw error;
  } finally {
    clearTimeout(timeoutId);
  }
}

function registryLabel(registry) {
  return registry === "mezzi" ? "Registro Mezzi" : "Registro Personale";
}

function isLegacyRegistryDeploymentError(error) {
  const message = String(error?.message || "");
  return /ID Google Sheet non valido|Azione non supportata|Intestazioni mancanti/i.test(message);
}

function legacyRegistryRows(payload) {
  const primary = payload.registry === "mezzi" ? "MEZZI" : "PERSONALE";
  const sourceRows = Array.isArray(payload.sheets?.[primary]) ? payload.sheets[primary] : [];
  const headerSet = new Set(["SYNC_KEY", "IMPIANTO_KEY"]);
  sourceRows.forEach((row) => Object.keys(row || {}).forEach((key) => headerSet.add(String(key))));
  const headers = Array.from(headerSet).slice(0, MAX_COLUMNS);
  const rows = sourceRows.map((row) => {
    const recordId = String(row?.RECORD_ID || row?.ID_OPERATORE || row?.ID_MEZZO || "");
    return headers.map((header) => {
      if (header === "SYNC_KEY" || header === "IMPIANTO_KEY") return recordId;
      return row?.[header] ?? "";
    });
  });
  return { primary, headers, rows };
}

async function callAppsScriptWithRegistryCompatibility(payload, user) {
  try {
    return { ...(await callAppsScript(payload, user)), legacyMode: false };
  } catch (error) {
    if (!isLegacyRegistryDeploymentError(error)) throw error;

    if (payload.action === "createRegistrySpreadsheet") {
      const result = await callAppsScript({
        action: "createSpreadsheet",
        commessaId: `registry-${payload.registry}`,
        commessaName: registryLabel(payload.registry),
        headers: [
          "SYNC_KEY",
          "IMPIANTO_KEY",
          "RECORD_ID",
          "UPDATED_AT",
          "UPDATED_BY",
          "SYNC_VERSION",
          "SYNC_SOURCE",
          "ROW_STATUS"
        ]
      }, user);
      return { ...result, legacyMode: true };
    }

    if (payload.action === "syncRegistrySpreadsheet") {
      const legacy = legacyRegistryRows(payload);
      const result = await callAppsScript({
        action: "replaceRows",
        spreadsheetId: payload.spreadsheetId,
        sheetUrl: payload.sheetUrl,
        gid: payload.gid || "0",
        headers: legacy.headers,
        rows: legacy.rows,
        commessaId: `registry-${payload.registry}`,
        commessaName: registryLabel(payload.registry),
        operationId: `registry-${payload.registry}-${Date.now()}`
      }, user);
      return {
        ...result,
        spreadsheetId: payload.spreadsheetId,
        sheetUrl: payload.sheetUrl,
        incoming: {},
        conflicts: [],
        rowsWritten: legacy.rows.length,
        legacyMode: true
      };
    }

    throw error;
  }
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return json(204, {});
  if (event.httpMethod !== "POST") return json(405, { ok: false, error: "Metodo non consentito." });
  try {
    const user = await authenticateEvent(event);
    const allowlist = allowedEmails();
    const email = String(user.email || "").toLowerCase();
    if (allowlist.size && !allowlist.has(email)) throw new Error("Utente non autorizzato alla scrittura del Google Sheet.");
    const payload = validatePayload(parseBody(event));
    const result = await callAppsScriptWithRegistryCompatibility(payload, user);
    return json(200, {
      ok: true,
      rowsWritten: Number(result.rowsWritten) || (Array.isArray(payload.rows) ? payload.rows.length : 0),
      sheetName: result.sheetName || "",
      spreadsheetId: result.spreadsheetId || payload.spreadsheetId,
      sheetUrl: result.sheetUrl || payload.sheetUrl || "",
      gid: String(result.gid ?? payload.gid ?? "0"),
      incoming: result.incoming || {},
      conflicts: result.conflicts || [],
      legacyMode: result.legacyMode === true
    });
  } catch (error) {
    console.error("Sincronizzazione Google Sheet non riuscita:", error);
    const message = error?.message || "Sincronizzazione Google Sheet non riuscita.";
    const status = /Token|Sessione|Utente non autorizzato/i.test(message) ? 401
      : /non configurato|APPS_SCRIPT_URL/i.test(message) ? 503
        : /non valid|non coincidono|Troppe righe|Intestazioni/i.test(message) ? 400 : 502;
    return json(status, { ok: false, error: message });
  }
};

exports.validatePayload = validatePayload;
exports.validateGoogleSheetUrl = validateGoogleSheetUrl;
exports.legacyRegistryRows = legacyRegistryRows;

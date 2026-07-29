"use strict";

const GOOGLE_SHEETS_HOST = "docs.google.com";
const FETCH_TIMEOUT_MS = 12000;
const MAX_CSV_BYTES = 8 * 1024 * 1024;

function parseGoogleSheetUrl(rawUrl) {
  let parsed;
  try {
    parsed = new URL(String(rawUrl || "").trim());
  } catch (_) {
    throw new Error("Link Google Sheet non valido.");
  }

  if (parsed.protocol !== "https:" || parsed.hostname !== GOOGLE_SHEETS_HOST) {
    throw new Error("Inserisci un link https://docs.google.com/spreadsheets valido.");
  }

  const match = parsed.pathname.match(/^\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  if (!match) throw new Error("Non riesco a trovare l'ID del foglio nel link.");

  const hashParams = new URLSearchParams(parsed.hash.replace(/^#/, ""));
  const requestedGid = parsed.searchParams.get("gid") || hashParams.get("gid") || "0";
  const gid = /^\d+$/.test(requestedGid) ? requestedGid : "0";
  return { spreadsheetId: match[1], gid };
}

function looksLikeHtml(text, contentType) {
  return /text\/html/i.test(contentType || "") || /^\s*(?:<!doctype html|<html\b)/i.test(text || "");
}

async function fetchCsv(spreadsheetId, gid) {
  const targets = [
    `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/gviz/tq?tqx=out:csv&gid=${encodeURIComponent(gid)}`,
    `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/export?format=csv&gid=${encodeURIComponent(gid)}`
  ];
  const errors = [];

  for (const target of targets) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);
    try {
      const response = await fetch(target, {
        redirect: "follow",
        signal: controller.signal,
        headers: {
          Accept: "text/csv,text/plain;q=0.9,*/*;q=0.1",
          "User-Agent": "VargaCantieri/1.0"
        }
      });
      const text = await response.text();
      const contentType = response.headers.get("content-type") || "";
      if (!response.ok) throw new Error(`Google HTTP ${response.status}`);
      if (!text.trim() || looksLikeHtml(text, contentType)) {
        throw new Error("foglio non pubblico o risposta Google non valida");
      }
      if (Buffer.byteLength(text, "utf8") > MAX_CSV_BYTES) {
        throw new Error("foglio troppo grande (massimo 8 MB)");
      }
      return text.replace(/^\uFEFF/, "");
    } catch (error) {
      errors.push(error?.name === "AbortError" ? "tempo di risposta scaduto" : (error?.message || String(error)));
    } finally {
      clearTimeout(timeoutId);
    }
  }

  throw new Error(errors.join("; "));
}

exports.handler = async (event) => {
  const corsHeaders = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Cache-Control": "no-store"
  };

  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers: corsHeaders, body: "" };
  if (event.httpMethod !== "GET") {
    return {
      statusCode: 405,
      headers: { ...corsHeaders, "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({ ok: false, error: "Metodo non consentito." })
    };
  }

  try {
    const { spreadsheetId, gid } = parseGoogleSheetUrl(event.queryStringParameters?.url);
    const csv = await fetchCsv(spreadsheetId, gid);
    return {
      statusCode: 200,
      headers: {
        ...corsHeaders,
        "Content-Type": "text/csv; charset=utf-8",
        "Content-Disposition": `inline; filename="google-sheet-${spreadsheetId}.csv"`
      },
      body: csv
    };
  } catch (error) {
    const message = error?.message || "Impossibile leggere il Google Sheet.";
    const invalidLink = /link|ID del foglio|docs\.google\.com/i.test(message);
    return {
      statusCode: invalidLink ? 400 : 502,
      headers: { ...corsHeaders, "Content-Type": "application/json; charset=utf-8" },
      body: JSON.stringify({
        ok: false,
        error: invalidLink
          ? message
          : "Google Sheet non leggibile. Imposta Condividi → Chiunque abbia il link → Visualizzatore e riprova.",
        detail: invalidLink ? "" : message
      })
    };
  }
};

exports.parseGoogleSheetUrl = parseGoogleSheetUrl;

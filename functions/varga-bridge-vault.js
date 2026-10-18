"use strict";

const { onCall, HttpsError } = require("firebase-functions/v2/https");
const { defineSecret } = require("firebase-functions/params");

const TOKEN = defineSecret("VARGA_MAP_TOKEN");
const URL = defineSecret("VARGA_MAP_BRIDGE_URL");
const OWNER_EMAIL = "ionut29019@gmail.com";
const ACTIONS = new Set([
  "ping", "configure", "installRequestSchedule", "requestScheduleStatus", "scan",
  "scanRequests", "receipts", "requestReceipts", "acknowledge",
  "acknowledgeRequests", "driveEnsureJob", "driveList", "driveCreateFolder",
  "driveUpload", "listDepurazioneConsuntivi", "saveDepurazioneConsuntivo",
  "getDepurazioneConsuntivo", "deleteDepurazioneConsuntivo", "driveGetFile",
  "driveRename", "driveTrash", "calendarUpsert", "calendarDelete"
]);

exports.callVargaBridgeVault = onCall({
  region: "europe-west1",
  secrets: [TOKEN, URL],
  timeoutSeconds: 60,
  memory: "512MiB"
}, async (request) => {
  if (!request.auth || request.auth.token.email_verified !== true ||
      String(request.auth.token.email || "").toLowerCase() !== OWNER_EMAIL) {
    throw new HttpsError("permission-denied", "Accesso alla cassaforte non autorizzato.");
  }
  const input = request.data;
  if (!input || typeof input !== "object" || Array.isArray(input) ||
      !ACTIONS.has(input.action) || Object.hasOwn(input, "token")) {
    throw new HttpsError("invalid-argument", "Richiesta al ponte non valida.");
  }
  const bridgeUrl = URL.value();
  const secret = TOKEN.value();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(bridgeUrl || "") || !secret) {
    throw new HttpsError("failed-precondition", "Cassaforte Drive da configurare sul server.");
  }
  const payload = JSON.stringify({ ...input, token: secret });
  if (Buffer.byteLength(payload) > 5 * 1024 * 1024) {
    throw new HttpsError("resource-exhausted", "Documento troppo grande per il ponte Drive.");
  }
  let response;
  try {
    response = await fetch(bridgeUrl, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: payload,
      signal: AbortSignal.timeout(55000)
    });
  } catch (_) {
    throw new HttpsError("unavailable", "Il ponte Drive non risponde.");
  }
  if (!response.ok) throw new HttpsError("unavailable", "Il ponte Drive ha restituito HTTP " + response.status);
  let result;
  try { result = await response.json(); } catch (_) {
    throw new HttpsError("internal", "Risposta Drive non valida.");
  }
  if (result?.ok === false && result.error === "Token ponte non valido") {
    throw new HttpsError("failed-precondition", "La chiave nella cassaforte non coincide con Apps Script.");
  }
  return result;
});

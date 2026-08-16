"use strict";

const crypto = require("node:crypto");
const functions = require("firebase-functions/v1");

const REGION = "europe-west1";
const ADMIN_EMAIL = "ionut29019@gmail.com";
const RESEND_ENDPOINT = "https://api.resend.com/emails";
const MAX_MESSAGE = 1200;
const MAX_STACK = 7000;
const MAX_TEXT = 900;
const RATE_WINDOW_MS = 10 * 60 * 1000;
const RATE_MAX_PER_USER = 8;
const recentByUser = new Map();

function cleanText(value, max = MAX_TEXT) {
  return String(value || "")
    .replace(/Bearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer [REDACTED]")
    .replace(/([?&](?:token|access_token|id_token|apikey|api_key|key|password|secret)=)[^&#\s]+/gi, "$1[REDACTED]")
    .replace(/\b(AIza[0-9A-Za-z_-]{20,})\b/g, "[REDACTED_API_KEY]")
    .slice(0, max);
}

function escapeHtml(value) {
  return cleanText(value, MAX_STACK)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function normalizeReport(data) {
  const raw = data && typeof data === "object" ? data : {};
  return {
    reportId: cleanText(raw.reportId, 100) || crypto.randomUUID(),
    fingerprint: cleanText(raw.fingerprint, 120),
    kind: cleanText(raw.kind, 80) || "runtime-error",
    message: cleanText(raw.message, MAX_MESSAGE) || "Errore senza messaggio",
    stack: cleanText(raw.stack, MAX_STACK),
    source: cleanText(raw.source, 700),
    line: Number.isFinite(Number(raw.line)) ? Number(raw.line) : null,
    column: Number.isFinite(Number(raw.column)) ? Number(raw.column) : null,
    occurredAt: cleanText(raw.occurredAt, 80) || new Date().toISOString(),
    page: cleanText(raw.page, 300),
    activeView: cleanText(raw.activeView, 160),
    online: raw.online !== false,
    visibility: cleanText(raw.visibility, 40),
    userAgent: cleanText(raw.userAgent, 900),
    platform: cleanText(raw.platform, 160),
    language: cleanText(raw.language, 40),
    screen: cleanText(raw.screen, 80),
    connection: cleanText(raw.connection, 160),
    appVersion: cleanText(raw.appVersion, 100)
  };
}

function diagnose(report) {
  const text = `${report.kind} ${report.message} ${report.stack}`.toLowerCase();
  if (/permission-denied|missing or insufficient permissions|unauthorized|forbidden/.test(text)) {
    return { category: "Permessi / autenticazione", severity: "alta", cause: "La richiesta è stata rifiutata per autorizzazioni o sessione non valida.", action: "Verificare sessione utente, ruolo e regole di accesso della funzione coinvolta." };
  }
  if (/failed to fetch|networkerror|network request failed|offline|internet disconnected|err_network/.test(text)) {
    return { category: "Connessione / rete", severity: "media", cause: "L'app non è riuscita a completare una richiesta di rete.", action: "Verificare connettività e disponibilità del servizio remoto; il client può riprovare quando torna online." };
  }
  if (/quota|resource-exhausted|too many requests|429/.test(text)) {
    return { category: "Quota / limite servizio", severity: "alta", cause: "Un servizio ha rifiutato la richiesta per limite o quota raggiunta.", action: "Controllare quote Firebase/API e verificare che non esistano richieste duplicate." };
  }
  if (/timeout|timed out|deadline-exceeded/.test(text)) {
    return { category: "Timeout / lentezza", severity: "media", cause: "Un'operazione ha superato il tempo massimo previsto.", action: "Controllare rete, backend e operazioni lente nella schermata indicata." };
  }
  if (/script error|loading chunk|failed to load|resource-error|404/.test(text)) {
    return { category: "Asset / PWA / cache", severity: "alta", cause: "Un file necessario all'app potrebbe non essere stato caricato correttamente.", action: "Verificare deploy, Service Worker, cache PWA e disponibilità del file indicato." };
  }
  if (/typeerror|referenceerror|cannot read|is not defined|undefined|null/.test(text)) {
    return { category: "Errore JavaScript", severity: "alta", cause: "Il codice ha incontrato uno stato o un dato inatteso durante l'esecuzione.", action: "Controllare stack, file e riga riportati, mantenendo invariati i flussi operativi non coinvolti." };
  }
  return { category: "Errore applicativo", severity: "media", cause: "Errore non gestito classificato automaticamente.", action: "Analizzare messaggio e stack tecnico per individuare il modulo responsabile." };
}

function enforceRateLimit(uid) {
  const now = Date.now();
  const recent = (recentByUser.get(uid) || []).filter((ts) => now - ts < RATE_WINDOW_MS);
  if (recent.length >= RATE_MAX_PER_USER) return false;
  recent.push(now);
  recentByUser.set(uid, recent);
  if (recentByUser.size > 1000) {
    for (const [key, times] of recentByUser.entries()) {
      if (!times.some((ts) => now - ts < RATE_WINDOW_MS)) recentByUser.delete(key);
    }
  }
  return true;
}

function buildEmail(report, user, diagnosis) {
  const location = [report.page, report.activeView].filter(Boolean).join(" · ") || "Non disponibile";
  const technicalLocation = [report.source, report.line ? `riga ${report.line}` : "", report.column ? `colonna ${report.column}` : ""].filter(Boolean).join(" · ") || "Non disponibile";
  const rows = [
    ["Utente", user.email || user.uid],
    ["UID", user.uid],
    ["Data/ora", report.occurredAt],
    ["Schermata", location],
    ["Tipo", report.kind],
    ["Diagnosi", diagnosis.category],
    ["Gravità", diagnosis.severity],
    ["Connessione", report.online ? `online${report.connection ? ` · ${report.connection}` : ""}` : "offline"],
    ["Dispositivo", report.platform || report.userAgent || "Non disponibile"],
    ["Schermo", report.screen || "Non disponibile"],
    ["Versione app", report.appVersion || "Non disponibile"],
    ["Posizione codice", technicalLocation],
    ["ID report", report.reportId],
    ["Fingerprint", report.fingerprint || "Non disponibile"]
  ];
  const text = [
    "VARGA CANTIERI - DIAGNOSTICA ERRORE",
    "",
    ...rows.map(([label, value]) => `${label}: ${value}`),
    "",
    `Problema: ${report.message}`,
    `Probabile causa: ${diagnosis.cause}`,
    `Azione consigliata: ${diagnosis.action}`,
    "",
    "STACK TECNICO:",
    report.stack || "Non disponibile"
  ].join("\n");
  const htmlRows = rows.map(([label, value]) => `<tr><td style="padding:6px 10px;color:#64748b;font-weight:700">${escapeHtml(label)}</td><td style="padding:6px 10px">${escapeHtml(value)}</td></tr>`).join("");
  const html = `<!doctype html><html><body style="font-family:Arial,sans-serif;color:#172033;background:#f4f7fb;padding:20px"><div style="max-width:760px;margin:auto;background:#fff;border-radius:16px;padding:22px"><h2 style="margin-top:0">🚨 Varga Cantieri · Errore rilevato</h2><table style="border-collapse:collapse;width:100%;background:#f8fafc;border-radius:12px">${htmlRows}</table><h3>Problema</h3><p>${escapeHtml(report.message)}</p><h3>Diagnosi automatica</h3><p><strong>${escapeHtml(diagnosis.category)}</strong> · gravità ${escapeHtml(diagnosis.severity)}</p><p>${escapeHtml(diagnosis.cause)}</p><p><strong>Azione consigliata:</strong> ${escapeHtml(diagnosis.action)}</p><h3>Stack tecnico</h3><pre style="white-space:pre-wrap;overflow-wrap:anywhere;background:#101827;color:#eef2ff;padding:14px;border-radius:10px;font-size:12px">${escapeHtml(report.stack || "Non disponibile")}</pre></div></body></html>`;
  return { text, html };
}

exports.reportClientError = functions
  .region(REGION)
  .runWith({ secrets: ["RESEND_API_KEY", "ERROR_REPORT_FROM"] })
  .https.onCall(async (data, context) => {
    if (!context.auth?.uid) {
      throw new functions.https.HttpsError("unauthenticated", "Accesso necessario per inviare la diagnostica.");
    }
    if (!enforceRateLimit(context.auth.uid)) {
      return { sent: false, rateLimited: true };
    }

    const apiKey = String(process.env.RESEND_API_KEY || "").trim();
    const from = String(process.env.ERROR_REPORT_FROM || "").trim();
    if (!apiKey || !from) {
      throw new functions.https.HttpsError("failed-precondition", "Sistema email diagnostica non configurato.");
    }

    const report = normalizeReport(data);
    const diagnosis = diagnose(report);
    const user = {
      uid: context.auth.uid,
      email: cleanText(context.auth.token?.email, 240),
      name: cleanText(context.auth.token?.name, 180)
    };
    const email = buildEmail(report, user, diagnosis);
    const idempotencyKey = `varga-error-${crypto.createHash("sha256").update(`${context.auth.uid}:${report.reportId}`).digest("hex").slice(0, 40)}`;

    let response;
    try {
      response = await fetch(RESEND_ENDPOINT, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
          "User-Agent": "varga-cantieri-error-reporter/1.0",
          "Idempotency-Key": idempotencyKey
        },
        body: JSON.stringify({
          from,
          to: [ADMIN_EMAIL],
          subject: `🚨 Varga Cantieri · ${diagnosis.category} · ${user.email || user.uid}`,
          text: email.text,
          html: email.html
        })
      });
    } catch (error) {
      console.error("Invio diagnostica email non raggiungibile.", { message: cleanText(error?.message, 300), reportId: report.reportId });
      throw new functions.https.HttpsError("unavailable", "Invio della diagnostica temporaneamente non disponibile.");
    }

    if (!response.ok) {
      const responseText = cleanText(await response.text().catch(() => ""), 700);
      console.error("Invio diagnostica email rifiutato.", { status: response.status, response: responseText, reportId: report.reportId });
      throw new functions.https.HttpsError("unavailable", "Il servizio email ha rifiutato la diagnostica.");
    }

    return { sent: true, reportId: report.reportId, diagnosis: { category: diagnosis.category, severity: diagnosis.severity } };
  });

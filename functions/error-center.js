"use strict";

const crypto = require("node:crypto");
const admin = require("firebase-admin");
const functions = require("firebase-functions/v1");

const REGION = "europe-west1";
const ADMIN_EMAIL = "ionut29019@gmail.com";
const GROUPS_COLLECTION = "appErrorGroups";
const SUMMARY_COLLECTION = "systemCounters";
const SUMMARY_DOCUMENT = "errorCenterSummary";
const RATE_WINDOW_MS = 10 * 60 * 1000;
const RATE_MAX_PER_USER = 40;
const DASHBOARD_READ_LIMIT = 200;
const PUBLIC_CALLABLE_OPTIONS = Object.freeze({ invoker: "public" });
const VALID_STATUSES = new Set(["open", "in_verification", "resolved", "ignored"]);
const VALID_SEVERITIES = new Set(["critical", "high", "medium", "low", "info"]);
const SEVERITY_RANK = { info: 0, low: 1, medium: 2, high: 3, critical: 4 };
const SENSITIVE_KEY = /password|passcode|pin|token|secret|cookie|authorization|credential|api.?key|session|jwt/i;
const recentByUser = new Map();

function db() {
  return admin.firestore();
}

function cleanText(value, max = 1800) {
  return String(value ?? "")
    .replace(/Bearer\s+[A-Za-z0-9._~+/=-]+/gi, "Bearer [RIMOSSO]")
    .replace(/([?&](?:token|access_token|id_token|apikey|api_key|key|password|secret|pin)=)[^&#\s]+/gi, "$1[RIMOSSO]")
    .replace(/\b(AIza[0-9A-Za-z_-]{20,})\b/g, "[CHIAVE_API_RIMOSSA]")
    .replace(/\beyJ[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\b/g, "[TOKEN_RIMOSSO]")
    .replace(/\b([A-Z0-9._%+-]{2})[A-Z0-9._%+-]*(@[A-Z0-9.-]+\.[A-Z]{2,})\b/gi, "$1***$2")
    .replace(/\b(password|passcode|pin|secret|token)\s*[:=]\s*[^\s,;]+/gi, "$1=[RIMOSSO]")
    .slice(0, max);
}

function sanitizeValue(value, depth = 0) {
  if (depth > 3 || value == null) return value == null ? null : "[DATI_RIDOTTI]";
  if (typeof value === "string") return cleanText(value, 1200);
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.slice(0, 20).map((item) => sanitizeValue(item, depth + 1));
  if (typeof value === "object") {
    const output = {};
    Object.entries(value).slice(0, 40).forEach(([key, item]) => {
      const safeKey = cleanText(key, 90);
      output[safeKey] = SENSITIVE_KEY.test(key) ? "[RIMOSSO]" : sanitizeValue(item, depth + 1);
    });
    return output;
  }
  return cleanText(value, 300);
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function maskEmail(value) {
  const email = normalizeEmail(value);
  const [local, domain] = email.split("@");
  if (!local || !domain) return cleanText(email, 180);
  return `${local.slice(0, 2)}***@${domain}`;
}

function userKey(uid) {
  return crypto.createHash("sha256").update(String(uid || "unknown")).digest("hex").slice(0, 16);
}

function safeFingerprint(value, fallbackSeed) {
  const raw = cleanText(value, 120).replace(/[^a-zA-Z0-9_-]/g, "");
  if (raw) return raw.slice(0, 120);
  return crypto.createHash("sha256").update(String(fallbackSeed || crypto.randomUUID())).digest("hex").slice(0, 40);
}

const FIRESTORE_STORAGE_QUOTA_FINGERPRINT = "firestore-local-storage-quota";

function isFirestoreStorageQuotaError(message, stack = "") {
  const text = `${message || ""}\n${stack || ""}`;
  return /(?:QuotaExceededError|exceeded the quota)/i.test(text)
    && /firestore_(?:clients|targets|mutations)_firestore/i.test(text);
}

function normalizeSeverity(value) {
  const severity = String(value || "").trim().toLowerCase();
  return VALID_SEVERITIES.has(severity) ? severity : "medium";
}

function maxSeverity(a, b) {
  return (SEVERITY_RANK[normalizeSeverity(a)] >= SEVERITY_RANK[normalizeSeverity(b)])
    ? normalizeSeverity(a)
    : normalizeSeverity(b);
}

function diagnose(report) {
  const text = `${report.kind} ${report.message} ${report.stack} ${report.feature}`.toLowerCase();
  let severity = normalizeSeverity(report.severity);
  let category = "Errore applicativo";
  let action = "Analizzare il dettaglio tecnico e le ultime azioni registrate.";

  if (isFirestoreStorageQuotaError(report.message, report.stack)) {
    category = "Archivio locale / Firestore";
    severity = maxSeverity(severity, "high");
    action = "Verificare il recupero automatico dello spazio locale e che la scrittura Firestore sia stata ritentata.";
  } else if (/ui-freeze|blocc|non risponde|freeze/.test(text) || Number(report.durationMs) >= 5000) {
    category = "Interfaccia bloccata";
    severity = "critical";
    action = "Controllare operazioni sincrone, observer e rendering avviati nella schermata indicata.";
  } else if (/repeated-tap|piu tocchi|più tocchi/.test(text)) {
    category = "Pulsante non reattivo";
    severity = maxSeverity(severity, "high");
    action = "Verificare overlay, pointer-events e lavoro eseguito sul thread principale dopo il primo tocco.";
  } else if (/permission-denied|missing or insufficient permissions|unauthorized|forbidden/.test(text)) {
    category = "Permessi / autenticazione";
    severity = maxSeverity(severity, "high");
    action = "Controllare sessione, ruolo utente, Cloud Function e regole di accesso.";
  } else if (/resource-error|failed to load|loading chunk|script error|404/.test(text)) {
    category = "Asset / PWA / cache";
    severity = maxSeverity(severity, "high");
    action = "Verificare deploy, Service Worker, cache e disponibilità del file indicato.";
  } else if (/typeerror|referenceerror|cannot read|is not defined|undefined|null|javascript-error|unhandled-rejection/.test(text)) {
    category = "Errore JavaScript";
    severity = maxSeverity(severity, "high");
    action = "Controllare stack, sorgente e stato dei dati nel momento dell'errore.";
  } else if (/timeout|timed out|deadline-exceeded|slow-interaction|long-task/.test(text) || Number(report.durationMs) >= 900) {
    category = "Lentezza / timeout";
    severity = maxSeverity(severity, Number(report.durationMs) >= 2500 ? "high" : "medium");
    action = "Ridurre il lavoro sincrono e verificare rete, API e caricamenti della funzione coinvolta.";
  } else if (/network|failed to fetch|offline|internet disconnected/.test(text)) {
    category = "Connessione / rete";
    severity = maxSeverity(severity, "medium");
    action = "Verificare connettività, disponibilità del servizio e gestione della coda offline.";
  } else if (report.manual) {
    category = "Segnalazione utente";
    severity = maxSeverity(severity, "medium");
    action = "Riprodurre i passaggi descritti dall'utente e confrontarli con le ultime azioni registrate.";
  }

  const featureLabel = cleanText(report.feature || report.activeView || "app", 100) || "app";
  const title = report.manual
    ? `Segnalazione utente · ${featureLabel}`
    : `${category} · ${featureLabel}`;
  return { severity, category, action, title };
}

function normalizeReport(data, context) {
  const raw = data && typeof data === "object" ? data : {};
  const report = {
    reportId: cleanText(raw.reportId, 100) || crypto.randomUUID(),
    kind: cleanText(raw.kind, 80) || "runtime-error",
    severity: normalizeSeverity(raw.severity),
    feature: cleanText(raw.feature, 120) || "app",
    message: cleanText(raw.message, 1800) || "Errore senza messaggio",
    stack: cleanText(raw.stack, 7000),
    source: cleanText(raw.source, 700),
    line: Number.isFinite(Number(raw.line)) ? Number(raw.line) : null,
    column: Number.isFinite(Number(raw.column)) ? Number(raw.column) : null,
    durationMs: Math.max(0, Math.min(120000, Math.round(Number(raw.durationMs || 0)))),
    tapCount: Math.max(0, Math.min(50, Math.round(Number(raw.tapCount || 0)))),
    manual: Boolean(raw.manual),
    occurredAt: cleanText(raw.occurredAt, 80) || new Date().toISOString(),
    page: cleanText(raw.page, 300),
    activeView: cleanText(raw.activeView, 160),
    online: raw.online !== false,
    visibility: cleanText(raw.visibility, 40),
    userAgent: cleanText(raw.userAgent, 900),
    platform: cleanText(raw.platform, 160),
    language: cleanText(raw.language, 40),
    screen: cleanText(raw.screen, 100),
    connection: cleanText(raw.connection, 160),
    appVersion: cleanText(raw.appVersion, 100),
    commessaId: cleanText(raw.commessaId, 160),
    commessaName: cleanText(raw.commessaName, 180),
    impiantoId: cleanText(raw.impiantoId, 160),
    metadata: sanitizeValue(raw.metadata || {}),
    breadcrumbs: Array.isArray(raw.breadcrumbs) ? raw.breadcrumbs.slice(-15).map((item) => sanitizeValue(item)) : []
  };
  report.fingerprint = !report.manual && isFirestoreStorageQuotaError(report.message, report.stack)
    ? FIRESTORE_STORAGE_QUOTA_FINGERPRINT
    : safeFingerprint(
      raw.fingerprint,
      `${report.kind}|${report.feature}|${report.message}|${report.stack.split("\n").slice(0, 2).join("|")}|${report.source}`
    );
  report.user = {
    uid: context.auth.uid,
    key: userKey(context.auth.uid),
    name: cleanText(context.auth.token?.name, 160),
    emailMasked: maskEmail(context.auth.token?.email)
  };
  report.diagnosis = diagnose(report);
  return report;
}

function enforceRateLimit(uid) {
  const now = Date.now();
  const recent = (recentByUser.get(uid) || []).filter((timestamp) => now - timestamp < RATE_WINDOW_MS);
  if (recent.length >= RATE_MAX_PER_USER) return false;
  recent.push(now);
  recentByUser.set(uid, recent);
  if (recentByUser.size > 1000) {
    for (const [key, times] of recentByUser.entries()) {
      if (!times.some((timestamp) => now - timestamp < RATE_WINDOW_MS)) recentByUser.delete(key);
    }
  }
  return true;
}

async function isAdminContext(context) {
  if (!context.auth?.uid) return false;
  const email = normalizeEmail(context.auth.token?.email);
  if (email === ADMIN_EMAIL || context.auth.token?.admin === true || context.auth.token?.isAdmin === true) return true;
  try {
    const snapshot = await db().collection("appConfig").doc("adminUsers").get();
    const emails = Array.isArray(snapshot.data()?.emails) ? snapshot.data().emails.map(normalizeEmail) : [];
    return emails.includes(email);
  } catch (_) {
    return false;
  }
}

async function requireAdmin(context) {
  if (!context.auth?.uid) throw new functions.https.HttpsError("unauthenticated", "Accesso richiesto.");
  if (!(await isAdminContext(context))) throw new functions.https.HttpsError("permission-denied", "Funzione riservata all'amministratore.");
}

function recentEvent(report) {
  return {
    reportId: report.reportId,
    occurredAt: report.occurredAt,
    kind: report.kind,
    severity: report.diagnosis.severity,
    message: report.message,
    page: report.page,
    activeView: report.activeView,
    source: report.source,
    line: report.line,
    column: report.column,
    durationMs: report.durationMs,
    tapCount: report.tapCount,
    platform: report.platform,
    appVersion: report.appVersion,
    online: report.online,
    connection: report.connection,
    commessaId: report.commessaId,
    commessaName: report.commessaName,
    impiantoId: report.impiantoId,
    metadata: report.metadata,
    breadcrumbs: report.breadcrumbs,
    userName: report.user.name,
    userEmailMasked: report.user.emailMasked
  };
}

function toIso(value) {
  try {
    if (value?.toDate) return value.toDate().toISOString();
    if (value instanceof Date) return value.toISOString();
    if (typeof value === "string") return value;
  } catch (_) {}
  return "";
}

function serializeGroup(id, data = {}) {
  return {
    id,
    fingerprint: cleanText(data.fingerprint || id, 120),
    title: cleanText(data.title, 200),
    category: cleanText(data.category, 140),
    feature: cleanText(data.feature, 140),
    severity: normalizeSeverity(data.severity),
    status: VALID_STATUSES.has(String(data.status || "")) ? String(data.status) : "open",
    occurrences: Math.max(0, Number(data.occurrences || 0)),
    affectedUsers: Math.max(0, Number(data.affectedUsers || 0)),
    manualCount: Math.max(0, Number(data.manualCount || 0)),
    firstSeenAt: toIso(data.firstSeenAt),
    lastSeenAt: toIso(data.lastSeenAt),
    lastMessage: cleanText(data.lastMessage, 1800),
    lastStack: cleanText(data.lastStack, 7000),
    lastSource: cleanText(data.lastSource, 700),
    lastPage: cleanText(data.lastPage, 300),
    lastActiveView: cleanText(data.lastActiveView, 160),
    lastPlatform: cleanText(data.lastPlatform, 160),
    lastAppVersion: cleanText(data.lastAppVersion, 100),
    lastDurationMs: Math.max(0, Number(data.lastDurationMs || 0)),
    lastTapCount: Math.max(0, Number(data.lastTapCount || 0)),
    lastUserName: cleanText(data.lastUserName, 160),
    lastUserEmailMasked: cleanText(data.lastUserEmailMasked, 180),
    commessaId: cleanText(data.commessaId, 160),
    commessaName: cleanText(data.commessaName, 180),
    impiantoId: cleanText(data.impiantoId, 160),
    diagnosisAction: cleanText(data.diagnosisAction, 1000),
    adminNote: cleanText(data.adminNote, 3000),
    updatedByName: cleanText(data.updatedByName, 160),
    updatedAt: toIso(data.updatedAt),
    recentEvents: Array.isArray(data.recentEvents) ? data.recentEvents.slice(0, 8).map((item) => sanitizeValue(item)) : [],
    statusHistory: Array.isArray(data.statusHistory) ? data.statusHistory.slice(-12).map((item) => sanitizeValue(item)) : []
  };
}

exports.recordClientErrorGroup = functions.region(REGION).runWith(PUBLIC_CALLABLE_OPTIONS).https.onCall(async (data, context) => {
  if (!context.auth?.uid) throw new functions.https.HttpsError("unauthenticated", "Accesso richiesto per registrare la diagnostica.");
  if (!enforceRateLimit(context.auth.uid)) return { recorded: false, rateLimited: true };

  const report = normalizeReport(data, context);
  const groupRef = db().collection(GROUPS_COLLECTION).doc(report.fingerprint);

  const transactionResult = await db().runTransaction(async (transaction) => {
    const snapshot = await transaction.get(groupRef);
    const previous = snapshot.exists ? snapshot.data() || {} : {};
    const previousStatus = VALID_STATUSES.has(String(previous.status || "")) ? String(previous.status) : "open";
    const reopened = ["resolved", "ignored"].includes(previousStatus)
      && (report.manual || ["critical", "high"].includes(report.diagnosis.severity));
    const users = Array.isArray(previous.affectedUserKeys) ? previous.affectedUserKeys.slice(0, 49) : [];
    if (!users.includes(report.user.key) && users.length < 50) users.push(report.user.key);
    const events = [recentEvent(report), ...(Array.isArray(previous.recentEvents) ? previous.recentEvents : [])].slice(0, 8);
    const occurrences = Math.max(0, Number(previous.occurrences || 0)) + 1;
    const severity = snapshot.exists ? maxSeverity(previous.severity, report.diagnosis.severity) : report.diagnosis.severity;

    transaction.set(groupRef, {
      fingerprint: report.fingerprint,
      title: report.diagnosis.title,
      category: report.diagnosis.category,
      feature: report.feature,
      severity,
      status: reopened ? "open" : previousStatus,
      occurrences,
      affectedUserKeys: users,
      affectedUsers: users.length,
      manualCount: Math.max(0, Number(previous.manualCount || 0)) + (report.manual ? 1 : 0),
      firstSeenAt: snapshot.exists && previous.firstSeenAt ? previous.firstSeenAt : admin.firestore.FieldValue.serverTimestamp(),
      lastSeenAt: admin.firestore.FieldValue.serverTimestamp(),
      lastMessage: report.message,
      lastStack: report.stack,
      lastSource: report.source,
      lastPage: report.page,
      lastActiveView: report.activeView,
      lastPlatform: report.platform,
      lastAppVersion: report.appVersion,
      lastDurationMs: report.durationMs,
      lastTapCount: report.tapCount,
      lastUserName: report.user.name,
      lastUserEmailMasked: report.user.emailMasked,
      commessaId: report.commessaId,
      commessaName: report.commessaName,
      impiantoId: report.impiantoId,
      diagnosisAction: report.diagnosis.action,
      recentEvents: events,
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });

    return { isNew: !snapshot.exists, reopened, occurrences, severity };
  });

  await db().collection(SUMMARY_COLLECTION).doc(SUMMARY_DOCUMENT).set({
    totalEvents: admin.firestore.FieldValue.increment(1),
    totalGroups: admin.firestore.FieldValue.increment(transactionResult.isNew ? 1 : 0),
    unseenAlerts: admin.firestore.FieldValue.increment(1),
    lastErrorAt: admin.firestore.FieldValue.serverTimestamp(),
    lastGroupId: report.fingerprint,
    lastSeverity: report.diagnosis.severity,
    lastTitle: report.diagnosis.title,
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });

  return {
    recorded: true,
    groupId: report.fingerprint,
    severity: report.diagnosis.severity,
    notified: false
  };
});

exports.getErrorCenterSummary = functions.region(REGION).runWith(PUBLIC_CALLABLE_OPTIONS).https.onCall(async (_data, context) => {
  await requireAdmin(context);
  const snapshot = await db().collection(SUMMARY_COLLECTION).doc(SUMMARY_DOCUMENT).get();
  const data = snapshot.exists ? snapshot.data() || {} : {};
  return {
    unseenAlerts: Math.max(0, Number(data.unseenAlerts || 0)),
    totalEvents: Math.max(0, Number(data.totalEvents || 0)),
    totalGroups: Math.max(0, Number(data.totalGroups || 0)),
    lastErrorAt: toIso(data.lastErrorAt),
    lastGroupId: cleanText(data.lastGroupId, 120),
    lastSeverity: normalizeSeverity(data.lastSeverity),
    lastTitle: cleanText(data.lastTitle, 200)
  };
});

exports.markErrorCenterSeen = functions.region(REGION).runWith(PUBLIC_CALLABLE_OPTIONS).https.onCall(async (_data, context) => {
  await requireAdmin(context);
  await db().collection(SUMMARY_COLLECTION).doc(SUMMARY_DOCUMENT).set({
    unseenAlerts: 0,
    lastSeenAt: admin.firestore.FieldValue.serverTimestamp(),
    lastSeenByUid: context.auth.uid,
    lastSeenByEmail: maskEmail(context.auth.token?.email),
    updatedAt: admin.firestore.FieldValue.serverTimestamp()
  }, { merge: true });
  return { updated: true };
});

exports.getErrorCenterDashboard = functions.region(REGION).runWith(PUBLIC_CALLABLE_OPTIONS).https.onCall(async (data, context) => {
  await requireAdmin(context);
  const input = data && typeof data === "object" ? data : {};
  const status = String(input.status || "all");
  const severity = String(input.severity || "all").toLowerCase();
  const queryText = cleanText(input.query, 160).toLowerCase();
  const requestedLimit = Math.max(1, Math.min(150, Number(input.limit || 100)));

  const groups = db().collection(GROUPS_COLLECTION);
  const [snapshot, totalCount, openCount, verificationCount, resolvedCount, ignoredCount, criticalCount, highCount] = await Promise.all([
    groups.orderBy("lastSeenAt", "desc").limit(DASHBOARD_READ_LIMIT + 1).get(),
    groups.count().get(),
    groups.where("status", "==", "open").count().get(),
    groups.where("status", "==", "in_verification").count().get(),
    groups.where("status", "==", "resolved").count().get(),
    groups.where("status", "==", "ignored").count().get(),
    groups.where("severity", "==", "critical").count().get(),
    groups.where("severity", "==", "high").count().get()
  ]);

  const all = snapshot.docs.slice(0, DASHBOARD_READ_LIMIT).map((doc) => serializeGroup(doc.id, doc.data() || {}));
  const items = all.filter((item) => {
    if (status !== "all" && item.status !== status) return false;
    if (severity !== "all" && item.severity !== severity) return false;
    if (queryText) {
      const haystack = `${item.title} ${item.category} ${item.feature} ${item.lastMessage} ${item.lastPage} ${item.commessaName}`.toLowerCase();
      if (!haystack.includes(queryText)) return false;
    }
    return true;
  }).slice(0, requestedLimit);

  const counts = {
    total: Math.max(0, Number(totalCount.data().count || 0)),
    open: Math.max(0, Number(openCount.data().count || 0)),
    inVerification: Math.max(0, Number(verificationCount.data().count || 0)),
    resolved: Math.max(0, Number(resolvedCount.data().count || 0)),
    ignored: Math.max(0, Number(ignoredCount.data().count || 0)),
    critical: Math.max(0, Number(criticalCount.data().count || 0)),
    high: Math.max(0, Number(highCount.data().count || 0))
  };

  return {
    items,
    counts,
    countsVerified: true,
    countSource: "firestore-aggregate",
    serverTime: new Date().toISOString(),
    truncated: snapshot.size > DASHBOARD_READ_LIMIT
  };
});

exports.updateErrorCenterStatus = functions.region(REGION).runWith(PUBLIC_CALLABLE_OPTIONS).https.onCall(async (data, context) => {
  await requireAdmin(context);
  const input = data && typeof data === "object" ? data : {};
  const groupId = safeFingerprint(input.groupId, "");
  const status = String(input.status || "").trim();
  const adminNote = cleanText(input.adminNote, 3000);
  if (!groupId) throw new functions.https.HttpsError("invalid-argument", "Gruppo errore non valido.");
  if (!VALID_STATUSES.has(status)) throw new functions.https.HttpsError("invalid-argument", "Stato non valido.");

  const ref = db().collection(GROUPS_COLLECTION).doc(groupId);
  await db().runTransaction(async (transaction) => {
    const snapshot = await transaction.get(ref);
    if (!snapshot.exists) throw new functions.https.HttpsError("not-found", "Errore non trovato.");
    const previous = snapshot.data() || {};
    const history = Array.isArray(previous.statusHistory) ? previous.statusHistory.slice(-11) : [];
    history.push({
      status,
      note: adminNote,
      at: new Date().toISOString(),
      byUid: context.auth.uid,
      byName: cleanText(context.auth.token?.name || context.auth.token?.email, 160)
    });
    transaction.set(ref, {
      status,
      adminNote,
      statusHistory: history,
      updatedByUid: context.auth.uid,
      updatedByName: cleanText(context.auth.token?.name || context.auth.token?.email, 160),
      updatedAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
  });

  return { updated: true, groupId, status };
});

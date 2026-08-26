"use strict";

const crypto = require("crypto");
const admin = require("firebase-admin");
const { onDocumentWritten } = require("firebase-functions/v2/firestore");

const REGION = "europe-west1";
const SHARED_COLLECTION = "sharedStaticViews";
const MAX_PAYLOAD_BYTES = 700000;
const CALENDAR_SCHEMA_VERSION = 2;
const SOURCE_COLLECTIONS = new Set(["oreReports", "oreApprovalRequests"]);
const ACTIVITY_COLLECTIONS = new Set(["impianti", "lavorazioni"]);
const COMPLETED_STATUSES = new Set(["fatto", "done", "completed", "completato"]);

function text(value) {
  return String(value ?? "").trim();
}

function dateFromValue(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value?.toDate === "function") {
    try {
      const converted = value.toDate();
      if (converted instanceof Date && !Number.isNaN(converted.getTime())) return converted;
    } catch (_) {}
  }
  if (typeof value === "number") {
    const converted = new Date(value);
    return Number.isNaN(converted.getTime()) ? null : converted;
  }
  const raw = text(value);
  if (!raw) return null;
  const italian = raw.match(/^(\d{2})[\/-](\d{2})[\/-](\d{4})/);
  if (italian) return new Date(`${italian[3]}-${italian[2]}-${italian[1]}T12:00:00Z`);
  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(`${iso[1]}-${iso[2]}-${iso[3]}T12:00:00Z`);
  const converted = new Date(raw);
  return Number.isNaN(converted.getTime()) ? null : converted;
}

function dateKeyFromData(data = {}) {
  const candidates = [
    data.date,
    data.data,
    data.giorno,
    data.workDate,
    data.reportDate,
    data.dataLavoro,
    data.createdForDate
  ];
  for (const candidate of candidates) {
    const raw = text(candidate);
    const direct = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (direct) return `${direct[1]}-${direct[2]}-${direct[3]}`;
    const parsed = dateFromValue(candidate);
    if (parsed) return parsed.toISOString().slice(0, 10);
  }
  return "";
}

function monthKeyFromData(data = {}) {
  const dateKey = dateKeyFromData(data);
  return /^\d{4}-\d{2}-\d{2}$/.test(dateKey) ? dateKey.slice(0, 7) : "";
}

function activityDateKeyFromData(data = {}) {
  return dateKeyFromData({
    date: data.dataEsecuzione || data.dataFatto || data.doneAt || data.completedAt || data.executionDate
  });
}

function normalizeStatus(value) {
  return text(value).toLowerCase();
}

function compactRecord(sourceCollection, sourceId, data = {}) {
  if (!SOURCE_COLLECTIONS.has(sourceCollection)) return null;
  const date = dateKeyFromData(data);
  if (!date) return null;
  const status = text(data.status || data.stato);
  if ([
    "rejected",
    "rifiutato",
    "rifiutata",
    "cancelled",
    "canceled",
    "annullato",
    "annullata"
  ].includes(normalizeStatus(status))) return null;

  // La vista condivisa deve conservare lo stesso schema letto dal calendario.
  // I campi originali restano intatti; vengono aggiunti soltanto i metadati
  // necessari per distinguere la raccolta sorgente e normalizzare la data.
  return {
    ...data,
    id: sourceId,
    sourceCollection,
    sourceKey: `${sourceCollection}/${sourceId}`,
    date,
    status
  };
}

function numberOrNull(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const raw = text(value).replace(/\s/g, "").replace(/\.(?=\d{3}(?:\D|$))/g, "").replace(",", ".").replace(/[^0-9.-]/g, "");
  if (!raw) return null;
  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed : null;
}

function completedAmount(data = {}) {
  for (const value of [data.totale, data.importo, data.totaleRiga, data.importoPrestazione, data.valoreProdotto, data.earnedAmount]) {
    const parsed = numberOrNull(value);
    if (parsed != null && parsed >= 0) return parsed;
  }
  return null;
}

function isCompletedActivity(sourceCollection, data = {}) {
  if (!ACTIVITY_COLLECTIONS.has(sourceCollection)) return false;
  const status = normalizeStatus(data.stato || data.status).replace(/_/g, " ");
  if (sourceCollection === "lavorazioni") return COMPLETED_STATUSES.has(status);
  return Boolean(data.done || data.fatto || data.completed || COMPLETED_STATUSES.has(status));
}

function compactActivity(sourceCollection, commessaId, itemId, data = {}) {
  if (!isCompletedActivity(sourceCollection, data)) return null;
  const date = activityDateKeyFromData(data);
  if (!date) return null;
  const sourceKey = `${sourceCollection}/${commessaId}/${itemId}`;
  const impiantoId = text(data.impiantoId || data.physicalPlantId || (sourceCollection === "impianti" ? itemId : ""));
  return {
    sourceCollection,
    sourceKey,
    kind: sourceCollection === "lavorazioni" ? "lavorazione" : "impianto",
    commessaId: text(commessaId || data.commessaId),
    commessaName: text(data.commessaName || data.commessaNome),
    itemId: text(itemId),
    impiantoId,
    idSap: text(data.idSap || data.idSAP || data.sapId),
    name: text(data.denominazione || data.nome || data.impiantoNome || (sourceCollection === "lavorazioni" ? "Lavorazione" : "Impianto")),
    comune: text(data.comune),
    address: text(data.indirizzo || data.descrizioneVia || data.via),
    work: text(data.tipologiaLavorazione || data.tipologiaIntervento || data.lavorazioniRichieste || data.codiceVocePrezzo || data.codicePrezzo),
    workCode: text(data.codiceVocePrezzo || data.codicePrezzo || data.voceRiferimento),
    operator: text(data.operatoreNome || data.operatore || data.doneBy || data.completedBy),
    operatorId: text(data.operatoreUid || data.doneByUid || data.userId),
    date,
    time: text(data.oraEsecuzione || data.oraFatto),
    amount: completedAmount(data),
    note: text(data.noteImpianto || data.note || data.nota),
    quantity: numberOrNull(data.quantita),
    unit: text(data.unitaMisura),
    economicStatus: text(data.economicStatus)
  };
}

function stableHash(payload) {
  return crypto.createHash("sha256").update(JSON.stringify(payload)).digest("hex");
}

function buildNextPayload(existingPayload, month, sourceCollection, sourceId, nextData) {
  const sourceKey = `${sourceCollection}/${sourceId}`;
  const current = Array.isArray(existingPayload?.reports) ? existingPayload.reports : [];
  const reports = current.filter((item) => text(item?.sourceKey || `${item?.sourceCollection || ""}/${item?.id || ""}`) !== sourceKey);
  const nextRecord = nextData ? compactRecord(sourceCollection, sourceId, nextData) : null;
  if (nextRecord && nextRecord.date.startsWith(month)) reports.push(nextRecord);
  reports.sort((a, b) => `${a.date}|${a.operatore || a.operatorName || ""}|${a.sourceKey}`.localeCompare(`${b.date}|${b.operatore || b.operatorName || ""}|${b.sourceKey}`, "it"));
  return {
    month,
    schemaVersion: CALENDAR_SCHEMA_VERSION,
    completeRecords: true,
    reports,
    activities: Array.isArray(existingPayload?.activities) ? existingPayload.activities : []
  };
}

function buildNextActivityPayload(existingPayload, month, sourceCollection, commessaId, itemId, nextData) {
  const sourceKey = `${sourceCollection}/${commessaId}/${itemId}`;
  const current = Array.isArray(existingPayload?.activities) ? existingPayload.activities : [];
  const activities = current.filter((item) => text(item?.sourceKey) !== sourceKey);
  const nextActivity = nextData ? compactActivity(sourceCollection, commessaId, itemId, nextData) : null;
  if (nextActivity && nextActivity.date.startsWith(month)) activities.push(nextActivity);
  activities.sort((a, b) => `${a.date}|${a.commessaId}|${a.impiantoId}|${a.sourceKey}`
    .localeCompare(`${b.date}|${b.commessaId}|${b.impiantoId}|${b.sourceKey}`, "it"));
  return {
    month,
    schemaVersion: CALENDAR_SCHEMA_VERSION,
    completeRecords: true,
    reports: Array.isArray(existingPayload?.reports) ? existingPayload.reports : [],
    activities
  };
}

async function updateMonthView({ month, sourceCollection, sourceId, commessaId = "", nextData, activity = false }) {
  if (!/^\d{4}-\d{2}$/.test(month)) return { skipped: "invalid-month" };
  const db = admin.firestore();
  const ref = db.collection(SHARED_COLLECTION).doc(`calendario__${month}`);

  return db.runTransaction(async (transaction) => {
    const snapshot = await transaction.get(ref);
    const previous = snapshot.exists ? snapshot.data() || {} : {};
    const payload = activity
      ? buildNextActivityPayload(previous.payload, month, sourceCollection, commessaId, sourceId, nextData)
      : buildNextPayload(previous.payload, month, sourceCollection, sourceId, nextData);
    const contentHash = stableHash(payload);
    if (previous.contentHash === contentHash) return { skipped: "unchanged" };

    const bytes = Buffer.byteLength(JSON.stringify(payload), "utf8");
    if (bytes > MAX_PAYLOAD_BYTES) {
      console.error("Vista calendario condivisa troppo grande; aggiornamento non pubblicato.", { month, bytes });
      return { skipped: "payload-too-large", bytes };
    }

    transaction.set(ref, {
      type: "calendario",
      key: month,
      version: Number(previous.version || 0) + 1,
      schemaVersion: CALENDAR_SCHEMA_VERSION,
      completeRecords: true,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAtClient: new Date().toISOString(),
      updatedBy: activity ? "cloud-function:activity-write" : "cloud-function:hours-write",
      authorName: activity ? "Aggiornamento automatico attività" : "Aggiornamento automatico ore",
      contentHash,
      payloadBytes: bytes,
      payload
    }, { merge: false });
    return { updated: true, month, reports: payload.reports.length, activities: payload.activities.length, bytes };
  });
}

async function handleActivityWrite(event, sourceCollection) {
  const beforeData = event.data?.before?.exists ? event.data.before.data() : null;
  const afterData = event.data?.after?.exists ? event.data.after.data() : null;
  const sourceId = event.params.itemId;
  const commessaId = event.params.commessaId;
  const beforeMonth = beforeData && isCompletedActivity(sourceCollection, beforeData)
    ? activityDateKeyFromData(beforeData).slice(0, 7)
    : "";
  const afterMonth = afterData && isCompletedActivity(sourceCollection, afterData)
    ? activityDateKeyFromData(afterData).slice(0, 7)
    : "";
  const months = [...new Set([beforeMonth, afterMonth].filter((month) => /^\d{4}-\d{2}$/.test(month)))];

  for (const month of months) {
    await updateMonthView({
      month,
      sourceCollection,
      sourceId,
      commessaId,
      nextData: afterMonth === month ? afterData : null,
      activity: true
    });
  }
  return null;
}

async function handleHoursWrite(event, sourceCollection) {
  const beforeData = event.data?.before?.exists ? event.data.before.data() : null;
  const afterData = event.data?.after?.exists ? event.data.after.data() : null;
  const sourceId = event.params.documentId;
  const beforeMonth = beforeData ? monthKeyFromData(beforeData) : "";
  const afterMonth = afterData ? monthKeyFromData(afterData) : "";
  const months = [...new Set([beforeMonth, afterMonth].filter(Boolean))];

  for (const month of months) {
    await updateMonthView({
      month,
      sourceCollection,
      sourceId,
      nextData: afterMonth === month ? afterData : null
    });
  }
  return null;
}

exports.syncSharedCalendarFromOreReports = onDocumentWritten(
  { document: "oreReports/{documentId}", region: REGION },
  (event) => handleHoursWrite(event, "oreReports")
);

exports.syncSharedCalendarFromOreApprovalRequests = onDocumentWritten(
  { document: "oreApprovalRequests/{documentId}", region: REGION },
  (event) => handleHoursWrite(event, "oreApprovalRequests")
);

exports.syncSharedCalendarFromImpianti = onDocumentWritten(
  { document: "commesse/{commessaId}/impianti/{itemId}", region: REGION },
  (event) => handleActivityWrite(event, "impianti")
);

exports.syncSharedCalendarFromLavorazioni = onDocumentWritten(
  { document: "commesse/{commessaId}/lavorazioni/{itemId}", region: REGION },
  (event) => handleActivityWrite(event, "lavorazioni")
);

exports.__test = {
  CALENDAR_SCHEMA_VERSION,
  dateKeyFromData,
  monthKeyFromData,
  activityDateKeyFromData,
  compactRecord,
  compactActivity,
  buildNextPayload,
  buildNextActivityPayload,
  stableHash
};

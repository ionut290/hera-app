"use strict";

const crypto = require("crypto");
const admin = require("firebase-admin");
const { onDocumentWritten } = require("firebase-functions/v2/firestore");

const REGION = "europe-west1";
const SHARED_COLLECTION = "sharedStaticViews";
const MAX_PAYLOAD_BYTES = 700000;
const SOURCE_COLLECTIONS = new Set(["oreReports", "oreApprovalRequests"]);

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

function normalizeStatus(value) {
  return text(value).toLowerCase();
}

function compactRecord(sourceCollection, sourceId, data = {}) {
  if (!SOURCE_COLLECTIONS.has(sourceCollection)) return null;
  const date = dateKeyFromData(data);
  if (!date) return null;
  const status = text(data.status || data.stato);
  if (["rejected", "rifiutato", "rifiutata", "annullato", "annullata"].includes(normalizeStatus(status))) return null;

  return {
    id: sourceId,
    sourceCollection,
    sourceKey: `${sourceCollection}/${sourceId}`,
    date,
    status,
    commessaId: text(data.commessaId || data.commessa || data.projectId),
    operatoreId: text(data.operatoreId || data.operatorId || data.userId || data.uid),
    operatore: text(data.operatore || data.operatorName || data.userName || data.nomeOperatore),
    entries: Array.isArray(data.entries) ? data.entries : [],
    updatedAtClient: text(data.updatedAtClient || data.modifiedAtClient || "")
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
  reports.sort((a, b) => `${a.date}|${a.operatore}|${a.sourceKey}`.localeCompare(`${b.date}|${b.operatore}|${b.sourceKey}`, "it"));
  return { month, reports };
}

async function updateMonthView({ month, sourceCollection, sourceId, nextData }) {
  if (!/^\d{4}-\d{2}$/.test(month)) return { skipped: "invalid-month" };
  const db = admin.firestore();
  const ref = db.collection(SHARED_COLLECTION).doc(`calendario__${month}`);

  return db.runTransaction(async (transaction) => {
    const snapshot = await transaction.get(ref);
    const previous = snapshot.exists ? snapshot.data() || {} : {};
    const payload = buildNextPayload(previous.payload, month, sourceCollection, sourceId, nextData);
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
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAtClient: new Date().toISOString(),
      updatedBy: "cloud-function:hours-write",
      authorName: "Aggiornamento automatico ore",
      contentHash,
      payload
    }, { merge: false });
    return { updated: true, month, reports: payload.reports.length };
  });
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

exports.__test = {
  dateKeyFromData,
  monthKeyFromData,
  compactRecord,
  buildNextPayload,
  stableHash
};

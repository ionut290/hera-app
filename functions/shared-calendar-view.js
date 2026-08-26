"use strict";

const crypto = require("crypto");
const admin = require("firebase-admin");
const { onDocumentWritten } = require("firebase-functions/v2/firestore");
const { onCall, HttpsError } = require("firebase-functions/v2/https");

const REGION = "europe-west1";
const SHARED_COLLECTION = "sharedStaticViews";
const MAX_PAYLOAD_BYTES = 700000;
const CALENDAR_SCHEMA_VERSION = 3;
const SOURCE_COLLECTIONS = new Set(["oreReports", "oreApprovalRequests"]);
const ACTIVITY_COLLECTIONS = new Set(["impianti", "lavorazioni"]);
const COMPLETED_STATUSES = new Set(["fatto", "done", "completed", "completato"]);
const MAX_RECOVERED_ACTIVITIES = 5000;

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
  const direct = dateKeyFromData({ date: data.dataEsecuzione || data.dataFatto || data.executionDate });
  if (direct) return direct;
  const timestamp = dateFromValue(data.doneAt || data.completedAt);
  if (!timestamp) return "";
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Europe/Rome",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(timestamp);
  const value = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  return `${value.year}-${value.month}-${value.day}`;
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
  let discountedPrice = null;
  for (const value of [data.prezzoRibassato, data.prezzoribassato, data.prezzoUnitarioRibassato, data.prezzounitarioribassato]) {
    const parsed = numberOrNull(value);
    if (parsed != null && parsed >= 0) {
      discountedPrice = parsed;
      break;
    }
  }
  if (discountedPrice == null) {
    const basePrice = numberOrNull(data.prezzoBase ?? data.prezzobase ?? data.prezzo);
    const rawDiscount = numberOrNull(data.percentualeRibasso);
    const discount = rawDiscount != null && rawDiscount > 1 ? rawDiscount / 100 : rawDiscount;
    if (basePrice != null && basePrice >= 0 && (discount == null || (discount >= 0 && discount <= 1))) {
      discountedPrice = basePrice * (1 - (discount ?? 0));
    }
  }
  if (discountedPrice == null) return null;
  const unit = text(data.unitaMisura || data.um).toUpperCase();
  if (unit === "AC") return discountedPrice;
  const quantity = numberOrNull(data.quantita);
  return quantity != null && quantity >= 0 ? quantity * discountedPrice : null;
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

function nextMonthKey(month) {
  const [year, monthNumber] = month.split("-").map(Number);
  const next = new Date(Date.UTC(year, monthNumber, 1));
  return `${next.getUTCFullYear()}-${String(next.getUTCMonth() + 1).padStart(2, "0")}`;
}

function activityCoordinatesFromSnapshot(snapshot, sourceCollection) {
  const segments = text(snapshot?.ref?.path).split("/");
  if (segments.length !== 4 || segments[0] !== "commesse" || segments[2] !== sourceCollection) return null;
  return { commessaId: segments[1], itemId: segments[3] };
}

async function recoverMonthActivities(month) {
  if (!/^\d{4}-\d{2}$/.test(month)) throw new HttpsError("invalid-argument", "Mese non valido.");
  const db = admin.firestore();
  const fromDate = `${month}-01`;
  const toDate = `${nextMonthKey(month)}-01`;
  const fromTimestamp = admin.firestore.Timestamp.fromMillis(Date.parse(`${fromDate}T00:00:00Z`) - 3 * 60 * 60 * 1000);
  const toTimestamp = admin.firestore.Timestamp.fromMillis(Date.parse(`${toDate}T00:00:00Z`) + 3 * 60 * 60 * 1000);
  const queryLimit = MAX_RECOVERED_ACTIVITIES + 1;
  const queries = [];

  for (const sourceCollection of ACTIVITY_COLLECTIONS) {
    const group = db.collectionGroup(sourceCollection);
    queries.push({
      sourceCollection,
      promise: group.where("dataEsecuzione", ">=", fromDate).where("dataEsecuzione", "<", toDate).limit(queryLimit).get()
    });
    queries.push({
      sourceCollection,
      promise: group.where("doneAt", ">=", fromTimestamp).where("doneAt", "<", toTimestamp).limit(queryLimit).get()
    });
  }

  const snapshots = await Promise.all(queries.map((query) => query.promise));
  const uniqueDocuments = new Map();
  snapshots.forEach((snapshot, index) => {
    const sourceCollection = queries[index].sourceCollection;
    snapshot.docs.forEach((document) => uniqueDocuments.set(document.ref.path, { document, sourceCollection }));
  });
  if (uniqueDocuments.size > MAX_RECOVERED_ACTIVITIES) {
    throw new HttpsError("resource-exhausted", "Troppe attività nel mese selezionato.");
  }

  const activities = [];
  uniqueDocuments.forEach(({ document, sourceCollection }) => {
    const coordinates = activityCoordinatesFromSnapshot(document, sourceCollection);
    if (!coordinates) return;
    const activity = compactActivity(sourceCollection, coordinates.commessaId, coordinates.itemId, document.data() || {});
    if (activity?.date.startsWith(month)) activities.push(activity);
  });
  activities.sort((a, b) => `${a.date}|${a.commessaId}|${a.impiantoId}|${a.sourceKey}`
    .localeCompare(`${b.date}|${b.commessaId}|${b.impiantoId}|${b.sourceKey}`, "it"));

  const ref = db.collection(SHARED_COLLECTION).doc(`calendario__${month}`);
  await db.runTransaction(async (transaction) => {
    const snapshot = await transaction.get(ref);
    const previous = snapshot.exists ? snapshot.data() || {} : {};
    const payload = {
      month,
      schemaVersion: CALENDAR_SCHEMA_VERSION,
      completeRecords: true,
      reports: Array.isArray(previous.payload?.reports) ? previous.payload.reports : [],
      activities
    };
    const contentHash = stableHash(payload);
    if (previous.contentHash === contentHash) return;
    const bytes = Buffer.byteLength(JSON.stringify(payload), "utf8");
    if (bytes > MAX_PAYLOAD_BYTES) throw new HttpsError("resource-exhausted", "Riepilogo mensile troppo grande.");
    transaction.set(ref, {
      type: "calendario",
      key: month,
      version: Number(previous.version || 0) + 1,
      schemaVersion: CALENDAR_SCHEMA_VERSION,
      completeRecords: true,
      updatedAt: admin.firestore.FieldValue.serverTimestamp(),
      updatedAtClient: new Date().toISOString(),
      updatedBy: "cloud-function:controlled-month-recovery",
      authorName: "Recupero mensile controllato",
      contentHash,
      payloadBytes: bytes,
      payload
    }, { merge: false });
  });

  return { month, schemaVersion: CALENDAR_SCHEMA_VERSION, activities, recovered: activities.length, complete: true };
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

exports.getAdministrativeCalendarMonth = onCall(
  { region: REGION, timeoutSeconds: 60, memory: "256MiB", invoker: "public" },
  async (request) => {
    if (!request.auth) throw new HttpsError("unauthenticated", "Login richiesto.");
    return recoverMonthActivities(text(request.data?.month));
  }
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
  nextMonthKey,
  activityCoordinatesFromSnapshot,
  stableHash
};

exports.__server = { recoverMonthActivities };

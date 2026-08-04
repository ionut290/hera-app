"use strict";

const admin = require("firebase-admin");
const { onDocumentWritten } = require("firebase-functions/v2/firestore");
const { onCall, HttpsError } = require("firebase-functions/v2/https");

const REGION = "europe-west1";
const SHARED_COLLECTION = "sharedStaticViews";
const MAX_PAYLOAD_BYTES = 700000;

const cleanRecord = (snapshot) => ({ id: snapshot.id, ...(snapshot.data() || {}) });

function dateKey(value) {
  if (!value) return "";
  if (typeof value?.toDate === "function") value = value.toDate();
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  const raw = String(value).trim();
  let match = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) return `${match[1]}-${match[2]}-${match[3]}`;
  match = raw.match(/^(\d{2})[/-](\d{2})[/-](\d{4})/);
  return match ? `${match[3]}-${match[2]}-${match[1]}` : "";
}

function dataDateKey(data = {}) {
  for (const value of [data.dateKey, data.date, data.data, data.giorno, data.workDate, data.selectedDate]) {
    const normalized = dateKey(value);
    if (normalized) return normalized;
  }
  return "";
}

async function writeSharedView(id, type, key, payload, updatedBy) {
  const bytes = Buffer.byteLength(JSON.stringify(payload), "utf8");
  if (bytes > MAX_PAYLOAD_BYTES) throw new Error(`Vista ${id} troppo grande: ${bytes} byte`);
  await admin.firestore().collection(SHARED_COLLECTION).doc(id).set({
    type, key, version: Date.now(),
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedBy, payload, payloadBytes: bytes
  }, { merge: false });
  return { id, count: Object.values(payload).find(Array.isArray)?.length || 0, bytes };
}

async function rebuildRegistryView() {
  const db = admin.firestore();
  const [personale, mezzi] = await Promise.all([db.collection("personale").get(), db.collection("mezzi").get()]);
  const result = await writeSharedView("registri__corrente", "registri", "corrente", {
    personale: personale.docs.map(cleanRecord), mezzi: mezzi.docs.map(cleanRecord)
  }, "cloud-function:registry-write");
  return { ...result, personale: personale.size, mezzi: mezzi.size };
}

async function readSquadreForDate(date) {
  const db = admin.firestore();
  const [historyByDateKey, historyByDate, current] = await Promise.all([
    db.collection("squadreStorico").where("dateKey", "==", date).get(),
    db.collection("squadreStorico").where("date", "==", date).get(),
    db.collection("squadreCommesse").get()
  ]);
  const unique = new Map();
  [...historyByDateKey.docs, ...historyByDate.docs].forEach((doc) => unique.set(doc.id, doc));
  current.docs.forEach((doc) => {
    const data = doc.data() || {};
    if (dataDateKey(data) === date) unique.set(doc.id, doc);
  });
  return [...unique.values()].map(cleanRecord);
}

async function rebuildSquadreDate(date) {
  const normalized = dateKey(date);
  if (!normalized) return null;
  const squadre = await readSquadreForDate(normalized);
  return writeSharedView(`squadre__${normalized}`, "squadre", normalized, { date: normalized, squadre }, "cloud-function:squadre-write");
}

async function rebuildSquadreFromWrite(event) {
  const before = event.data?.before?.exists ? dataDateKey(event.data.before.data() || {}) : "";
  const after = event.data?.after?.exists ? dataDateKey(event.data.after.data() || {}) : "";
  const dates = [...new Set([before, after].filter(Boolean))];
  return Promise.all(dates.map(rebuildSquadreDate));
}

async function rebuildCalendarMonth(month) {
  if (!/^\d{4}-\d{2}$/.test(month)) return null;
  const db = admin.firestore();
  const from = `${month}-01`;
  const to = `${month}-31`;
  const [reports, approvals] = await Promise.all([
    db.collection("oreReports").where("date", ">=", from).where("date", "<=", to).get(),
    db.collection("oreApprovalRequests").where("date", ">=", from).where("date", "<=", to).get()
  ]);
  const rows = [
    ...reports.docs.map((doc) => ({ ...cleanRecord(doc), sourceCollection: "oreReports" })),
    ...approvals.docs.map((doc) => ({ ...cleanRecord(doc), sourceCollection: "oreApprovalRequests" }))
      .filter((row) => !["rejected", "rifiutato", "rifiutata", "cancelled", "annullato", "annullata"].includes(String(row.status || row.stato || "").toLowerCase()))
  ];
  return writeSharedView(`calendario__${month}`, "calendario", month, { month, reports: rows }, "callable:shared-view-backfill");
}

async function assertAdmin(request) {
  if (!request.auth) throw new HttpsError("unauthenticated", "Autenticazione richiesta.");
  if (request.auth.token?.admin === true) return;
  const email = String(request.auth.token?.email || "").toLowerCase();
  const config = await admin.firestore().collection("appConfig").doc("adminUsers").get();
  const data = config.data() || {};
  const configured = [data.emails, data.adminEmails, data.users].flat().filter(Boolean).map((value) => String(value).toLowerCase());
  if (email === "ionut29019@gmail.com" || configured.includes(email)) return;
  throw new HttpsError("permission-denied", "Permessi amministratore richiesti.");
}

exports.rebuildAllSharedStaticViews = onCall({ region: REGION, timeoutSeconds: 300, invoker: "public" }, async (request) => {
  await assertAdmin(request);
  const requestedDate = dateKey(request.data?.date || "2026-08-04") || "2026-08-04";
  const requestedMonth = String(request.data?.month || requestedDate.slice(0, 7));
  const [registri, squadre, calendario] = await Promise.all([
    rebuildRegistryView(), rebuildSquadreDate(requestedDate), rebuildCalendarMonth(requestedMonth)
  ]);
  const result = { registri, squadre, calendario, completedAt: new Date().toISOString() };
  console.log("Backfill sharedStaticViews completato", result);
  return result;
});

exports.syncSharedRegistriesFromPersonale = onDocumentWritten({ document: "personale/{documentId}", region: REGION }, rebuildRegistryView);
exports.syncSharedRegistriesFromMezzi = onDocumentWritten({ document: "mezzi/{documentId}", region: REGION }, rebuildRegistryView);
exports.syncSharedSquadreFromHistory = onDocumentWritten({ document: "squadreStorico/{documentId}", region: REGION }, rebuildSquadreFromWrite);
exports.syncSharedSquadreFromCurrent = onDocumentWritten({ document: "squadreCommesse/{documentId}", region: REGION }, rebuildSquadreFromWrite);

exports.__test = { dateKey, dataDateKey };

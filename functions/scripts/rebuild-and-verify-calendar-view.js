"use strict";

const admin = require("firebase-admin");

const PROJECT_ID = process.env.GCLOUD_PROJECT || "hera-app-6cd2b";
const MONTH = process.env.CALENDAR_MONTH || "2026-08";
const SHARED_COLLECTION = "sharedStaticViews";
const MAX_PAYLOAD_BYTES = 700000;
const EXCLUDED = new Set(["rejected", "rifiutato", "rifiutata", "cancelled", "canceled", "annullato", "annullata"]);

if (!admin.apps.length) admin.initializeApp({ projectId: PROJECT_ID });
const db = admin.firestore();

function cleanRecord(doc, sourceCollection) {
  return {
    id: doc.id,
    sourceCollection,
    sourceKey: `${sourceCollection}/${doc.id}`,
    ...(doc.data() || {})
  };
}

function normalizedStatus(row) {
  return String(row.status || row.stato || "").trim().toLowerCase();
}

async function readMonth(collectionName) {
  const from = `${MONTH}-01`;
  const to = `${MONTH}-31`;
  const snapshot = await db.collection(collectionName)
    .where("date", ">=", from)
    .where("date", "<=", to)
    .get();
  return snapshot.docs.map((doc) => cleanRecord(doc, collectionName));
}

function sortedKeys(rows) {
  return rows.map((row) => row.sourceKey).sort();
}

function sameKeys(left, right) {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

async function main() {
  if (!/^\d{4}-\d{2}$/.test(MONTH)) throw new Error(`Mese non valido: ${MONTH}`);

  const ref = db.collection(SHARED_COLLECTION).doc(`calendario__${MONTH}`);
  const [reports, approvalsRaw, existing] = await Promise.all([
    readMonth("oreReports"),
    readMonth("oreApprovalRequests"),
    ref.get()
  ]);
  const approvals = approvalsRaw.filter((row) => !EXCLUDED.has(normalizedStatus(row)));
  const rows = [...reports, ...approvals];
  const payload = {
    month: MONTH,
    schemaVersion: 2,
    completeRecords: true,
    reports: rows,
    activities: Array.isArray(existing.data()?.payload?.activities) ? existing.data().payload.activities : []
  };
  const bytes = Buffer.byteLength(JSON.stringify(payload), "utf8");
  if (bytes > MAX_PAYLOAD_BYTES) {
    throw new Error(`Vista calendario troppo grande: ${bytes} byte`);
  }

  await ref.set({
    type: "calendario",
    key: MONTH,
    schemaVersion: 2,
    completeRecords: true,
    version: Date.now(),
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedAtClient: new Date().toISOString(),
    updatedBy: "github-action:calendar-safe-backfill",
    payload,
    payloadBytes: bytes
  }, { merge: false });

  const written = await ref.get();
  const writtenRows = written.data()?.payload?.reports || [];
  const expectedKeys = sortedKeys(rows);
  const writtenKeys = sortedKeys(writtenRows);

  if (!sameKeys(expectedKeys, writtenKeys)) {
    const missing = expectedKeys.filter((key) => !writtenKeys.includes(key));
    const extra = writtenKeys.filter((key) => !expectedKeys.includes(key));
    throw new Error(`Verifica fallita. Mancanti: ${missing.join(", ") || "nessuno"}; extra: ${extra.join(", ") || "nessuno"}`);
  }

  if (written.data()?.schemaVersion !== 2 || written.data()?.completeRecords !== true) {
    throw new Error("Verifica fallita: marcatori di completezza assenti.");
  }

  console.log(JSON.stringify({
    success: true,
    month: MONTH,
    oreReports: reports.length,
    oreApprovalRequests: approvals.length,
    total: rows.length,
    payloadBytes: bytes,
    verifiedKeys: writtenKeys.length
  }, null, 2));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

"use strict";

const admin = require("firebase-admin");
const { onDocumentWritten } = require("firebase-functions/v2/firestore");

const REGION = "europe-west1";
const SHARED_COLLECTION = "sharedStaticViews";
const MAX_PAYLOAD_BYTES = 700000;

function cleanRecord(snapshot) {
  return { id: snapshot.id, ...(snapshot.data() || {}) };
}

async function writeSharedView(id, type, key, payload, updatedBy) {
  const bytes = Buffer.byteLength(JSON.stringify(payload), "utf8");
  if (bytes > MAX_PAYLOAD_BYTES) throw new Error(`Vista ${id} troppo grande: ${bytes} byte`);
  await admin.firestore().collection(SHARED_COLLECTION).doc(id).set({
    type, key,
    version: admin.firestore.FieldValue.increment(1),
    updatedAt: admin.firestore.FieldValue.serverTimestamp(),
    updatedBy,
    payload
  }, { merge: true });
}

async function rebuildRegistryView() {
  const db = admin.firestore();
  const [personale, mezzi] = await Promise.all([
    db.collection("personale").get(),
    db.collection("mezzi").get()
  ]);
  return writeSharedView("registri__corrente", "registri", "corrente", {
    personale: personale.docs.map(cleanRecord),
    mezzi: mezzi.docs.map(cleanRecord)
  }, "cloud-function:registry-write");
}

async function rebuildSquadreView(event) {
  const data = event.data?.after?.exists ? event.data.after.data() : event.data?.before?.data();
  const date = String(data?.dateKey || "").slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return null;
  const snapshot = await admin.firestore().collection("squadreStorico").where("dateKey", "==", date).get();
  return writeSharedView(`squadre__${date}`, "squadre", date, {
    date,
    squadre: snapshot.docs.map(cleanRecord)
  }, "cloud-function:squadre-write");
}

exports.syncSharedRegistriesFromPersonale = onDocumentWritten(
  { document: "personale/{documentId}", region: REGION },
  rebuildRegistryView
);
exports.syncSharedRegistriesFromMezzi = onDocumentWritten(
  { document: "mezzi/{documentId}", region: REGION },
  rebuildRegistryView
);
exports.syncSharedSquadreFromHistory = onDocumentWritten(
  { document: "squadreStorico/{documentId}", region: REGION },
  rebuildSquadreView
);

exports.__test = { cleanRecord };

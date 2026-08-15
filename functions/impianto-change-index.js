"use strict";

const { onDocumentWritten } = require("firebase-functions/v2/firestore");
const admin = require("firebase-admin");

if (!admin.apps.length) admin.initializeApp();

const REGION = "europe-west1";

function getChangeType(beforeExists, afterExists) {
  if (!beforeExists && afterExists) return "created";
  if (beforeExists && !afterExists) return "deleted";
  return "updated";
}

exports.syncImpiantoChangeIndex = onDocumentWritten(
  {
    document: "commesse/{commessaId}/impianti/{impiantoId}",
    region: REGION
  },
  async (event) => {
    const commessaId = String(event.params?.commessaId || "").trim();
    const impiantoId = String(event.params?.impiantoId || "").trim();
    if (!commessaId || !impiantoId) return;

    const before = event.data?.before;
    const after = event.data?.after;
    const beforeExists = Boolean(before?.exists);
    const afterExists = Boolean(after?.exists);
    const source = afterExists ? (after.data() || {}) : (before?.data() || {});

    const markerRef = admin
      .firestore()
      .collection("commesse")
      .doc(commessaId)
      .collection("impiantoChangeIndex")
      .doc(impiantoId);

    await markerRef.set({
      commessaId,
      impiantoId,
      deleted: !afterExists,
      changeType: getChangeType(beforeExists, afterExists),
      changedAt: admin.firestore.FieldValue.serverTimestamp(),
      sourceUpdatedAt: source.updatedAt || null,
      sourceDoneAt: source.doneAt || null
    });
  }
);

exports.__test = { getChangeType };

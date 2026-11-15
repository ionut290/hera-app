"use strict";

const { onDocumentWritten } = require("firebase-functions/v2/firestore");
const admin = require("firebase-admin");

if (!admin.apps.length) admin.initializeApp();

const REGION = "europe-west1";
const SYNC_MARKER = "canonical-coordinate-sync-v1";

function numberOrNull(value) {
  if (value === null || value === undefined || String(value).trim() === "") return null;
  const number = Number(String(value).replace(",", "."));
  return Number.isFinite(number) ? number : null;
}

function canonicalCoordinates(data = {}) {
  const latitude = numberOrNull(data.latitudine ?? data.gpsY);
  const longitude = numberOrNull(data.longitudine ?? data.gpsX);
  return { latitudine: latitude, longitudine: longitude, gpsY: latitude, gpsX: longitude };
}

function coordinatePatch(data = {}, source) {
  return {
    ...canonicalCoordinates(data),
    coordinateSyncOrigin: SYNC_MARKER,
    coordinateSyncSource: source,
    coordinateSyncedAt: admin.firestore.FieldValue.serverTimestamp()
  };
}

function hasSameCoordinates(left = {}, right = {}) {
  const a = canonicalCoordinates(left);
  const b = canonicalCoordinates(right);
  return a.latitudine === b.latitudine && a.longitudine === b.longitudine;
}

async function synchronizeCounterpart(event, sourceCollection, targetCollection) {
  const after = event.data?.after;
  if (!after?.exists) return null;

  const data = after.data() || {};
  if (data.coordinateSyncOrigin === SYNC_MARKER) return null;

  const commessaId = String(event.params?.commessaId || "").trim();
  const impiantoId = String(event.params?.impiantoId || "").trim();
  if (!commessaId || !impiantoId) return null;

  const db = admin.firestore();
  const targetRef = db.collection("commesse").doc(commessaId).collection(targetCollection).doc(impiantoId);
  const target = await targetRef.get();
  if (target.exists && hasSameCoordinates(data, target.data() || {})) return null;

  await targetRef.set(coordinatePatch(data, sourceCollection), { merge: true });
  return null;
}

exports.syncPhysicalPlantCoordinates = onDocumentWritten(
  { document: "commesse/{commessaId}/impiantiFisici/{impiantoId}", region: REGION },
  (event) => synchronizeCounterpart(event, "impiantiFisici", "impianti")
);

exports.syncOperationalPlantCoordinates = onDocumentWritten(
  { document: "commesse/{commessaId}/impianti/{impiantoId}", region: REGION },
  (event) => synchronizeCounterpart(event, "impianti", "impiantiFisici")
);

exports.__test = { canonicalCoordinates, hasSameCoordinates };

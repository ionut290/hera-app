"use strict";

const { onDocumentWritten } = require("firebase-functions/v2/firestore");
const admin = require("firebase-admin");

if (!admin.apps.length) admin.initializeApp();

const REGION = "europe-west1";
const SYNC_MARKER = "canonical-coordinate-sync-v2";
const COORDINATE_FIELDS = ["latitudine", "longitudine", "gpsY", "gpsX"];

function numberOrNull(value) {
  if (value === null || value === undefined || String(value).trim() === "") return null;
  const number = Number(String(value).replace(",", "."));
  return Number.isFinite(number) ? number : null;
}

function preferredCoordinate(data, previous, primaryKey, aliasKey) {
  // Map editors may change only gpsX/gpsY while table editors change latitudine/longitudine.
  // Respect whichever coordinate field the user actually edited.
  if (previous) {
    const primaryChanged = numberOrNull(data[primaryKey]) !== numberOrNull(previous[primaryKey]);
    const aliasChanged = numberOrNull(data[aliasKey]) !== numberOrNull(previous[aliasKey]);
    if (primaryChanged && numberOrNull(data[primaryKey]) !== null) return numberOrNull(data[primaryKey]);
    if (aliasChanged && numberOrNull(data[aliasKey]) !== null) return numberOrNull(data[aliasKey]);
  }
  return numberOrNull(data[primaryKey]) ?? numberOrNull(data[aliasKey]);
}

function canonicalCoordinates(data = {}, previous = null) {
  const latitude = preferredCoordinate(data, previous, "latitudine", "gpsY");
  const longitude = preferredCoordinate(data, previous, "longitudine", "gpsX");
  return { latitudine: latitude, longitudine: longitude, gpsY: latitude, gpsX: longitude };
}

function validCoordinates(data) {
  const { latitudine: lat, longitudine: lon } = canonicalCoordinates(data);
  return lat !== null && lon !== null && lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180;
}

function coordinateFieldsChanged(before = {}, after = {}) {
  return COORDINATE_FIELDS.some((field) => numberOrNull(before[field]) !== numberOrNull(after[field]));
}

function stamp(value) {
  if (!value) return "";
  if (typeof value.toMillis === "function") return String(value.toMillis());
  if (typeof value.seconds === "number") return String(value.seconds) + ":" + String(value.nanoseconds || 0);
  return String(value);
}

function isFunctionGeneratedWrite(before = {}, after = {}) {
  return after.coordinateSyncOrigin === SYNC_MARKER &&
    !!stamp(after.coordinateSyncedAt) &&
    stamp(after.coordinateSyncedAt) !== stamp(before.coordinateSyncedAt);
}

function hasSameCoordinates(canonical, target = {}) {
  return COORDINATE_FIELDS.every((key) => numberOrNull(target[key]) === canonical[key]);
}

function coordinatePatch(canonical, source) {
  return {
    ...canonical,
    coordinateSyncOrigin: SYNC_MARKER,
    coordinateSyncSource: source,
    coordinateSyncedAt: admin.firestore.FieldValue.serverTimestamp()
  };
}

async function synchronizeCounterpart(event, sourceCollection) {
  const after = event.data?.after;
  if (!after?.exists) return null;

  const data = after.data() || {};
  const before = event.data?.before?.exists ? event.data.before.data() || {} : null;
  if (before && !coordinateFieldsChanged(before, data)) return null;
  if (isFunctionGeneratedWrite(before || {}, data)) return null;

  const canonical = canonicalCoordinates(data, before);
  if (!validCoordinates(canonical)) return null; // Never replace good GPS with missing/invalid GPS.

  const commessaId = String(event.params?.commessaId || "").trim();
  const impiantoId = String(event.params?.impiantoId || "").trim();
  if (!commessaId || !impiantoId) return null;

  const root = admin.firestore().collection("commesse").doc(commessaId);
  const targets = new Map();
  if (sourceCollection === "impianti") {
    // Operational records may use an ID different from their linked physical plant.
    const physicalId = String(data.physicalPlantId || impiantoId).trim();
    if (!physicalId) return null;
    const ref = root.collection("impiantiFisici").doc(physicalId);
    const snapshot = await ref.get();
    if (snapshot.exists) targets.set(ref.path, snapshot);
  } else {
    const operational = root.collection("impianti");
    const [sameId, linked] = await Promise.all([
      operational.doc(impiantoId).get(),
      operational.where("physicalPlantId", "==", impiantoId).get()
    ]);
    if (sameId.exists) targets.set(sameId.ref.path, sameId);
    linked.docs.forEach((snapshot) => targets.set(snapshot.ref.path, snapshot));
  }

  for (const snapshot of targets.values()) {
    if (hasSameCoordinates(canonical, snapshot.data() || {})) continue;
    await snapshot.ref.set(coordinatePatch(canonical, sourceCollection), { merge: true });
  }
  return null;
}

exports.syncPhysicalPlantCoordinates = onDocumentWritten(
  { document: "commesse/{commessaId}/impiantiFisici/{impiantoId}", region: REGION },
  (event) => synchronizeCounterpart(event, "impiantiFisici")
);

exports.syncOperationalPlantCoordinates = onDocumentWritten(
  { document: "commesse/{commessaId}/impianti/{impiantoId}", region: REGION },
  (event) => synchronizeCounterpart(event, "impianti")
);

exports.__test = {
  numberOrNull, canonicalCoordinates, validCoordinates,
  coordinateFieldsChanged, isFunctionGeneratedWrite, hasSameCoordinates
};

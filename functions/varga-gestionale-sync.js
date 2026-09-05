"use strict";

const crypto = require("node:crypto");
const admin = require("firebase-admin");
const { onCall, HttpsError } = require("firebase-functions/v2/https");
const { onSchedule } = require("firebase-functions/v2/scheduler");
const { onDocumentWritten } = require("firebase-functions/v2/firestore");

const REGION = "europe-west1";
const SNAPSHOT_COLLECTION = "gestionaleSyncSnapshots";
const CHANGE_INDEX_COLLECTION = "gestionaleSyncChangeIndex";
const STATE_COLLECTION = "gestionaleSyncState";
const STATE_DOC = "current";
const SCHEMA_VERSION = 2;
const CHUNK_TARGET_BYTES = 320000;
const MAX_RECURSION_DEPTH = 3;
const CHANGE_PAGE_LIMIT = 200;
const MAX_CHANGE_PAGE_LIMIT = 400;

const ROOT_SOURCES = [
  { name: "commesse", recursive: true },
  { name: "impianti" },
  { name: "squadre" },
  { name: "squadreCommesse" },
  { name: "squadreStorico" },
  { name: "oreReports" },
  { name: "oreApprovalRequests" },
  { name: "personale" },
  { name: "mezzi" },
  { name: "documents", filter: (data) => String(data?.visibility || "").toLowerCase() !== "personal" },
  { name: "posDocuments" },
  { name: "calendarEvents" },
  { name: "programmazioni" },
  { name: "verdeLevatoCommesse" },
  { name: "verdeLevatoRecords" },
  { name: "servizioNeveClienti" },
  { name: "servizioNevePercorsi" },
  { name: "servizioNeveMezzi" },
  { name: "servizioNeveOperatori" },
  { name: "servizioNeveSegnalazioni" },
  { name: "neve_commesse", recursive: true },
  { name: "neve_squadre" },
  { name: "neve_squadre_storico" }
];

const BLOCKED_COLLECTION_NAMES = new Set([
  "privateDocuments",
  "activityLogs",
  "auditLogsWhazzup",
  "operatorPositions",
  "fcmTokens",
  "notifications",
  "passwordRequests",
  "driveAdminSecret"
]);

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

async function isAdminRequest(request) {
  if (!request.auth) return false;
  if (request.auth.token?.admin === true) return true;
  const email = normalizeEmail(request.auth.token?.email);
  if (email === "ionut29019@gmail.com") return true;
  const config = await admin.firestore().collection("appConfig").doc("adminUsers").get();
  const data = config.data() || {};
  const configured = [data.emails, data.adminEmails, data.users]
    .flat()
    .filter(Boolean)
    .map(normalizeEmail);
  return configured.includes(email);
}

async function assertAdmin(request) {
  if (!request.auth) throw new HttpsError("unauthenticated", "Autenticazione richiesta.");
  if (!(await isAdminRequest(request))) {
    throw new HttpsError("permission-denied", "Permessi amministratore richiesti.");
  }
}

async function assertGestionaleReader(request) {
  if (!request.auth) throw new HttpsError("unauthenticated", "Autenticazione richiesta.");
  if (await isAdminRequest(request)) return;
  const email = normalizeEmail(request.auth.token?.email);
  const uid = String(request.auth.uid || "");
  const config = await admin.firestore().collection("appConfig").doc("gestionaleUsers").get();
  const data = config.data() || {};
  const emails = [data.emails, data.users, data.allowedEmails].flat().filter(Boolean).map(normalizeEmail);
  const uids = [data.uids, data.userIds, data.allowedUids].flat().filter(Boolean).map(String);
  if (emails.includes(email) || uids.includes(uid)) return;
  throw new HttpsError("permission-denied", "Utente non autorizzato a Varga Gestionale.");
}

function isGeoPointValue(value) {
  return Boolean(
    value &&
    typeof value === "object" &&
    typeof value.latitude === "number" &&
    typeof value.longitude === "number" &&
    value.constructor?.name === "GeoPoint"
  );
}

function isDocumentReferenceValue(value) {
  return Boolean(
    value &&
    typeof value === "object" &&
    typeof value.path === "string" &&
    value.firestore &&
    value.constructor?.name === "DocumentReference"
  );
}

function serializeValue(value) {
  if (value == null) return value;
  if (typeof value?.toDate === "function") {
    const date = value.toDate();
    return Number.isNaN(date?.getTime?.()) ? null : date.toISOString();
  }
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value.toISOString();
  if (isGeoPointValue(value)) return { latitude: value.latitude, longitude: value.longitude };
  if (isDocumentReferenceValue(value)) return { path: value.path };
  if (Array.isArray(value)) return value.map(serializeValue);
  if (typeof value === "object") {
    const output = {};
    Object.entries(value).forEach(([key, item]) => {
      if (typeof item === "undefined" || typeof item === "function") return;
      output[key] = serializeValue(item);
    });
    return output;
  }
  return value;
}

function recordFromDoc(doc, rootCollection) {
  return {
    sourcePath: doc.ref.path,
    rootCollection,
    id: doc.id,
    data: serializeValue(doc.data() || {})
  };
}

function collectionNameFromPath(path) {
  const parts = String(path || "").split("/").filter(Boolean);
  return parts.length ? parts[parts.length - 1] : "";
}

function shouldSkipCollection(path) {
  const name = collectionNameFromPath(path);
  return BLOCKED_COLLECTION_NAMES.has(name) || /^appConfig$/i.test(name);
}

function sourceDefinitionForPath(path) {
  const parts = String(path || "").split("/").filter(Boolean);
  if (!parts.length || parts.length % 2 !== 0) return null;
  const source = ROOT_SOURCES.find((item) => item.name === parts[0]);
  if (!source) return null;
  const depth = (parts.length / 2) - 1;
  if (depth > 0 && !source.recursive) return null;
  if (depth > MAX_RECURSION_DEPTH) return null;
  for (let index = 0; index < parts.length; index += 2) {
    if (shouldSkipCollection(parts[index])) return null;
  }
  return source;
}

function sourceAllowsData(source, data) {
  return Boolean(source && (!source.filter || source.filter(data || {})));
}

function changeDocumentId(sourcePath) {
  return crypto.createHash("sha256").update(String(sourcePath || "")).digest("hex");
}

function timestampParts(value, fallbackIso) {
  if (value && Number.isFinite(Number(value.seconds))) {
    return {
      seconds: Number(value.seconds),
      nanoseconds: Number(value.nanoseconds || 0)
    };
  }
  const millis = Date.parse(String(fallbackIso || ""));
  const safeMillis = Number.isFinite(millis) ? millis : Date.now();
  return {
    seconds: Math.floor(safeMillis / 1000),
    nanoseconds: (safeMillis % 1000) * 1000000
  };
}

function changeVersion(value, sourcePath = "", fallbackIso = "") {
  const parts = timestampParts(value, fallbackIso);
  const seconds = String(parts.seconds).padStart(12, "0");
  const nanoseconds = String(parts.nanoseconds).padStart(9, "0");
  const suffix = sourcePath ? changeDocumentId(sourcePath).slice(0, 16) : "";
  return `${seconds}.${nanoseconds}:${suffix}`;
}

function trackedChangeOperation(source, before, after) {
  const beforeIncluded = Boolean(before?.exists && sourceAllowsData(source, before.data()));
  const afterIncluded = Boolean(after?.exists && sourceAllowsData(source, after.data()));
  if (!beforeIncluded && !afterIncluded) return "";
  return afterIncluded ? "upsert" : "delete";
}

async function queueVargaGestionaleChange(event) {
  const before = event.data?.before;
  const after = event.data?.after;
  const snapshot = after?.exists ? after : before;
  const sourcePath = String(snapshot?.ref?.path || "");
  const source = sourceDefinitionForPath(sourcePath);
  if (!source) return null;
  const operation = trackedChangeOperation(source, before, after);
  if (!operation) return null;

  const sourceTimestamp = (after?.exists ? after.updateTime : before?.updateTime) || null;
  const sourceVersion = changeVersion(sourceTimestamp, sourcePath, event.time);
  const ref = admin.firestore().collection(CHANGE_INDEX_COLLECTION).doc(changeDocumentId(sourcePath));
  await admin.firestore().runTransaction(async (transaction) => {
    const current = await transaction.get(ref);
    if (current.exists && String(current.data()?.sourceVersion || "") >= sourceVersion) return;
    const version = changeVersion(admin.firestore.Timestamp.now(), sourcePath);
    transaction.set(ref, {
      sourcePath,
      rootCollection: source.name,
      sourceId: String(snapshot?.id || ""),
      operation,
      version,
      sourceVersion,
      eventId: String(event.id || ""),
      changedAt: admin.firestore.FieldValue.serverTimestamp()
    }, { merge: false });
  });
  return { sourcePath, operation, sourceVersion };
}

async function readCollection(collectionRef, rootCollection, options = {}, depth = 0) {
  if (shouldSkipCollection(collectionRef.path)) return [];
  const snapshot = await collectionRef.get();
  const records = [];
  for (const doc of snapshot.docs) {
    const data = doc.data() || {};
    if (!options.filter || options.filter(data, doc)) records.push(recordFromDoc(doc, rootCollection));
    if (options.recursive && depth < MAX_RECURSION_DEPTH) {
      const nested = await doc.ref.listCollections();
      for (const child of nested) {
        if (shouldSkipCollection(child.path)) continue;
        const childRecords = await readCollection(child, rootCollection, { recursive: true }, depth + 1);
        records.push(...childRecords);
      }
    }
  }
  return records;
}

function chunkRecords(records) {
  const chunks = [];
  let current = [];
  let bytes = 2;
  for (const record of records) {
    const recordBytes = Buffer.byteLength(JSON.stringify(record), "utf8") + 1;
    if (current.length && bytes + recordBytes > CHUNK_TARGET_BYTES) {
      chunks.push(current);
      current = [];
      bytes = 2;
    }
    current.push(record);
    bytes += recordBytes;
  }
  if (current.length) chunks.push(current);
  return chunks;
}

function countByRoot(records) {
  return records.reduce((counts, record) => {
    counts[record.rootCollection] = (counts[record.rootCollection] || 0) + 1;
    return counts;
  }, {});
}

async function deleteSnapshot(snapshotId) {
  if (!snapshotId) return;
  const db = admin.firestore();
  const chunks = await db.collection(SNAPSHOT_COLLECTION).doc(snapshotId).collection("chunks").get();
  for (let i = 0; i < chunks.docs.length; i += 400) {
    const batch = db.batch();
    chunks.docs.slice(i, i + 400).forEach((doc) => batch.delete(doc.ref));
    await batch.commit();
  }
  await db.collection(SNAPSHOT_COLLECTION).doc(snapshotId).delete().catch(() => undefined);
}

async function buildSnapshot() {
  const db = admin.firestore();
  const baselineTimestamp = admin.firestore.Timestamp.now();
  const changeCursor = changeVersion(baselineTimestamp);
  const records = [];
  for (const source of ROOT_SOURCES) {
    try {
      console.log("Varga Gestionale snapshot: leggo", source.name);
      const sourceRecords = await readCollection(db.collection(source.name), source.name, source);
      records.push(...sourceRecords);
      console.log("Varga Gestionale snapshot: completata", source.name, sourceRecords.length);
    } catch (error) {
      console.error("Varga Gestionale snapshot: errore raccolta", source.name, error);
      throw new HttpsError(
        "internal",
        `Errore durante la lettura della raccolta ${source.name}: ${String(error?.message || error)}`
      );
    }
  }

  records.sort((a, b) => String(a.sourcePath).localeCompare(String(b.sourcePath)));
  const chunks = chunkRecords(records);
  const snapshotId = `sync_${Date.now()}`;
  const generatedAt = new Date().toISOString();
  const manifest = {
    schemaVersion: SCHEMA_VERSION,
    snapshotId,
    generatedAt,
    totalRecords: records.length,
    chunkCount: chunks.length,
    counts: countByRoot(records),
    source: "Varga Cantieri",
    mode: "full-snapshot",
    changeCursor,
    incremental: true,
    originalIdsPreserved: true
  };

  const snapshotRef = db.collection(SNAPSHOT_COLLECTION).doc(snapshotId);
  await snapshotRef.set({ ...manifest, status: "building" });
  for (let i = 0; i < chunks.length; i += 200) {
    const batch = db.batch();
    chunks.slice(i, i + 200).forEach((rows, offset) => {
      const index = i + offset;
      batch.set(snapshotRef.collection("chunks").doc(String(index).padStart(5, "0")), {
        index,
        records: rows,
        recordCount: rows.length,
        payloadBytes: Buffer.byteLength(JSON.stringify(rows), "utf8")
      });
    });
    await batch.commit();
  }
  await snapshotRef.set({ ...manifest, status: "ready" }, { merge: true });

  const stateRef = db.collection(STATE_COLLECTION).doc(STATE_DOC);
  const previous = await stateRef.get();
  const previousSnapshotId = previous.data()?.snapshotId || "";
  await stateRef.set({ ...manifest, updatedAt: admin.firestore.FieldValue.serverTimestamp() }, { merge: false });
  if (previousSnapshotId && previousSnapshotId !== snapshotId) await deleteSnapshot(previousSnapshotId);
  return manifest;
}

async function ensureIncrementalBaseline() {
  const stateRef = admin.firestore().collection(STATE_COLLECTION).doc(STATE_DOC);
  const current = await stateRef.get();
  const data = current.data() || {};
  if (current.exists && Number(data.schemaVersion) >= SCHEMA_VERSION && data.changeCursor) {
    return serializeValue(data);
  }
  return buildSnapshot();
}

function deletionRecord(change) {
  const sourcePath = String(change.sourcePath || "");
  const parts = sourcePath.split("/").filter(Boolean);
  return {
    sourcePath,
    rootCollection: String(change.rootCollection || parts[0] || ""),
    id: String(change.sourceId || parts[parts.length - 1] || ""),
    operation: "delete",
    deleted: true,
    deletedAt: String(change.version || ""),
    data: null
  };
}

async function readIncrementalChanges(cursor, requestedLimit) {
  const limit = Math.min(
    MAX_CHANGE_PAGE_LIMIT,
    Math.max(1, Number(requestedLimit) || CHANGE_PAGE_LIMIT)
  );
  const snapshot = await admin.firestore()
    .collection(CHANGE_INDEX_COLLECTION)
    .where("version", ">", cursor)
    .orderBy("version")
    .limit(limit)
    .get();
  const changes = snapshot.docs.map((doc) => doc.data() || {});
  const upsertRefs = changes
    .filter((change) => String(change.operation || "") !== "delete")
    .filter((change) => sourceDefinitionForPath(change.sourcePath))
    .map((change) => admin.firestore().doc(String(change.sourcePath)));
  const currentDocs = upsertRefs.length ? await admin.firestore().getAll(...upsertRefs) : [];
  const currentByPath = new Map(currentDocs.map((doc) => [doc.ref.path, doc]));
  const records = changes.map((change) => {
    const sourcePath = String(change.sourcePath || "");
    const source = sourceDefinitionForPath(sourcePath);
    if (!source) return null;
    const current = currentByPath.get(sourcePath);
    if (
      String(change.operation || "") === "delete" ||
      !current?.exists ||
      !sourceAllowsData(source, current.data())
    ) {
      return deletionRecord(change);
    }
    return { ...recordFromDoc(current, source.name), operation: "upsert" };
  }).filter(Boolean);
  const nextCursor = snapshot.docs.length
    ? String(snapshot.docs[snapshot.docs.length - 1].data()?.version || cursor)
    : cursor;
  return {
    schemaVersion: SCHEMA_VERSION,
    mode: "incremental",
    records,
    recordCount: records.length,
    nextCursor,
    hasMore: snapshot.docs.length === limit
  };
}

function triggerPathForSource(sourceName, depth) {
  let path = `${sourceName}/{documentId0}`;
  for (let index = 1; index <= depth; index += 1) {
    path += `/{collectionId${index}}/{documentId${index}}`;
  }
  return path;
}

function triggerKeyForSource(sourceName, depth) {
  const normalized = String(sourceName || "")
    .replace(/[^a-zA-Z0-9]+(.)/g, (_, character) => String(character || "").toUpperCase());
  return `${normalized}D${depth}`;
}

const incrementalChangeTriggers = {};
ROOT_SOURCES.forEach((source) => {
  const maxDepth = source.recursive ? MAX_RECURSION_DEPTH : 0;
  for (let depth = 0; depth <= maxDepth; depth += 1) {
    incrementalChangeTriggers[triggerKeyForSource(source.name, depth)] = onDocumentWritten({
      document: triggerPathForSource(source.name, depth),
      region: REGION
    }, queueVargaGestionaleChange);
  }
});
exports.vgDelta = incrementalChangeTriggers;

exports.rebuildVargaGestionaleSnapshot = onCall({
  region: REGION,
  timeoutSeconds: 540,
  memory: "1GiB",
  invoker: "public",
  cors: [
    "https://ionut290.github.io",
    "https://creative-syrniki-dddbae.netlify.app",
    "http://localhost:3000",
    "http://localhost:5173",
    "http://localhost:8080"
  ]
}, async (request) => {
  await assertAdmin(request);
  const manifest = await buildSnapshot();
  console.log("Snapshot Varga Gestionale completato", manifest);
  return manifest;
});

const scheduledSyncOptions = {
  region: REGION,
  timeZone: "Europe/Rome",
  timeoutSeconds: 540,
  memory: "1GiB",
  retryCount: 1
};

async function runScheduledVargaGestionaleSync(timeLabel) {
  const manifest = await ensureIncrementalBaseline();
  await admin.firestore().collection(STATE_COLLECTION).doc(STATE_DOC).set({
    lastScheduledAt: admin.firestore.FieldValue.serverTimestamp(),
    lastScheduledSlot: timeLabel,
    incremental: true
  }, { merge: true });
  console.log(`Sincronizzazione incrementale Varga Cantieri pronta (${timeLabel})`, {
    snapshotId: manifest.snapshotId,
    changeCursor: manifest.changeCursor
  });
  return {
    incremental: true,
    snapshotId: manifest.snapshotId,
    changeCursor: manifest.changeCursor,
    scheduledSlot: timeLabel
  };
}

exports.syncVargaGestionaleAt0530 = onSchedule({
  ...scheduledSyncOptions,
  schedule: "30 5 * * *"
}, async () => {
  return runScheduledVargaGestionaleSync("05:30");
});

exports.syncVargaGestionaleAt1600 = onSchedule({
  ...scheduledSyncOptions,
  schedule: "0 16 * * *"
}, async () => {
  return runScheduledVargaGestionaleSync("16:00");
});

exports.getVargaGestionaleSnapshotManifest = onCall({
  region: REGION,
  timeoutSeconds: 60,
  invoker: "public",
  cors: true
}, async (request) => {
  await assertGestionaleReader(request);
  const state = await admin.firestore().collection(STATE_COLLECTION).doc(STATE_DOC).get();
  if (!state.exists) return { available: false };
  return { available: true, ...(serializeValue(state.data() || {})) };
});

exports.getVargaGestionaleSnapshotChunk = onCall({
  region: REGION,
  timeoutSeconds: 60,
  invoker: "public",
  cors: true
}, async (request) => {
  await assertGestionaleReader(request);
  const snapshotId = String(request.data?.snapshotId || "").trim();
  const index = Number(request.data?.index);
  if (!snapshotId || !Number.isInteger(index) || index < 0) {
    throw new HttpsError("invalid-argument", "snapshotId e index validi sono obbligatori.");
  }
  const ref = admin.firestore().collection(SNAPSHOT_COLLECTION).doc(snapshotId).collection("chunks").doc(String(index).padStart(5, "0"));
  const chunk = await ref.get();
  if (!chunk.exists) throw new HttpsError("not-found", "Blocco di sincronizzazione non trovato.");
  return serializeValue(chunk.data() || {});
});

exports.getVargaGestionaleChanges = onCall({
  region: REGION,
  timeoutSeconds: 120,
  memory: "512MiB",
  invoker: "public",
  cors: true
}, async (request) => {
  await assertGestionaleReader(request);
  const cursor = String(request.data?.cursor || "").trim();
  if (!cursor) {
    const manifest = await ensureIncrementalBaseline();
    return {
      schemaVersion: SCHEMA_VERSION,
      mode: "bootstrap",
      bootstrapRequired: true,
      manifest: serializeValue(manifest)
    };
  }
  return readIncrementalChanges(cursor, request.data?.limit);
});

exports.__test = {
  SCHEMA_VERSION,
  ROOT_SOURCES: ROOT_SOURCES.map((source) => source.name),
  chunkRecords,
  serializeValue,
  countByRoot,
  sourceDefinitionForPath,
  changeVersion,
  trackedChangeOperation,
  triggerPathForSource,
  triggerKeyForSource,
  deletionRecord
};

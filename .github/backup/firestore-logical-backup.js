#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");

const admin = require(path.join(process.cwd(), "functions", "node_modules", "firebase-admin"));

const projectId = String(process.env.PROJECT_ID || "hera-app-6cd2b").trim();
const outputDir = path.resolve(process.argv[2] || path.join(process.cwd(), "firestore-logical-backup"));
const MAX_COLLECTION_WORKERS = Math.max(2, Number(process.env.FIRESTORE_BACKUP_COLLECTION_WORKERS || 12));
const MAX_CHILD_LIST_CALLS = Math.max(4, Number(process.env.FIRESTORE_BACKUP_CHILD_WORKERS || 32));
const PAGE_SIZE = Math.max(50, Math.min(500, Number(process.env.FIRESTORE_BACKUP_PAGE_SIZE || 250)));
fs.mkdirSync(outputDir, { recursive: true });

if (!admin.apps.length) {
  admin.initializeApp({
    credential: admin.credential.applicationDefault(),
    projectId
  });
}
const db = admin.firestore();

const Timestamp = admin.firestore.Timestamp;
const GeoPoint = admin.firestore.GeoPoint;
const FieldPath = admin.firestore.FieldPath;

function encode(value) {
  if (value === null) return { t: "null" };
  if (value === undefined) return { t: "undefined" };
  if (typeof value === "string") return { t: "string", v: value };
  if (typeof value === "boolean") return { t: "boolean", v: value };
  if (typeof value === "bigint") return { t: "bigint", v: value.toString() };
  if (typeof value === "number") {
    if (Number.isNaN(value)) return { t: "number-special", v: "NaN" };
    if (value === Infinity) return { t: "number-special", v: "Infinity" };
    if (value === -Infinity) return { t: "number-special", v: "-Infinity" };
    return { t: "number", v: value };
  }
  if (value instanceof Timestamp || (value && typeof value.toMillis === "function" && Number.isFinite(Number(value.seconds)))) {
    return { t: "timestamp", s: Number(value.seconds), n: Number(value.nanoseconds || 0) };
  }
  if (value instanceof GeoPoint || (
    value && value.constructor?.name === "GeoPoint"
    && Number.isFinite(Number(value.latitude)) && Number.isFinite(Number(value.longitude))
  )) {
    return { t: "geopoint", lat: Number(value.latitude), lng: Number(value.longitude) };
  }
  if (value && typeof value.path === "string" && value.firestore && value.constructor?.name === "DocumentReference") {
    return { t: "reference", path: value.path };
  }
  if (Buffer.isBuffer(value)) return { t: "bytes", b64: value.toString("base64") };
  if (value instanceof Uint8Array) return { t: "bytes", b64: Buffer.from(value).toString("base64") };
  if (value instanceof Date) return { t: "date", iso: value.toISOString() };
  if (Array.isArray(value)) return { t: "array", v: value.map(encode) };
  if (typeof value === "object") {
    const mapped = {};
    for (const [key, entry] of Object.entries(value)) mapped[key] = encode(entry);
    return { t: "map", v: mapped };
  }
  throw new Error(`Tipo Firestore non serializzabile: ${typeof value}`);
}

function timestampIso(value) {
  try {
    if (!value) return null;
    if (typeof value.toDate === "function") return value.toDate().toISOString();
    if (value instanceof Date) return value.toISOString();
  } catch (_) {}
  return null;
}

function createLimiter(limit) {
  let active = 0;
  const queue = [];
  const runNext = () => {
    while (active < limit && queue.length) {
      const task = queue.shift();
      active += 1;
      Promise.resolve()
        .then(task.fn)
        .then(task.resolve, task.reject)
        .finally(() => {
          active -= 1;
          runNext();
        });
    }
  };
  return (fn) => new Promise((resolve, reject) => {
    queue.push({ fn, resolve, reject });
    runNext();
  });
}

async function main() {
  const startedAt = new Date().toISOString();
  const jsonlPath = path.join(outputDir, "documents.jsonl");
  const manifestPath = path.join(outputDir, "manifest.json");
  const out = fs.createWriteStream(jsonlPath, { encoding: "utf8", flags: "w" });
  const seenDocuments = new Set();
  const seenCollections = new Set();
  const topLevelCollections = await db.listCollections();
  const childListLimit = createLimiter(MAX_CHILD_LIST_CALLS);

  let documentCount = 0;
  let collectionCount = 0;
  let maxDepth = 0;
  let writeChain = Promise.resolve();
  let fatalError = null;

  out.on("error", (error) => {
    fatalError = fatalError || error;
  });

  function writeRecord(record) {
    const line = `${JSON.stringify(record)}\n`;
    writeChain = writeChain.then(() => new Promise((resolve, reject) => {
      if (fatalError) return reject(fatalError);
      if (out.write(line)) return resolve();
      const cleanup = () => {
        out.off("drain", onDrain);
        out.off("error", onError);
      };
      const onDrain = () => {
        cleanup();
        resolve();
      };
      const onError = (error) => {
        cleanup();
        reject(error);
      };
      out.once("drain", onDrain);
      out.once("error", onError);
    }));
    return writeChain;
  }

  const collectionQueue = [];
  let activeCollections = 0;
  let doneResolve;
  let doneReject;
  let settled = false;
  const allCollectionsDone = new Promise((resolve, reject) => {
    doneResolve = resolve;
    doneReject = reject;
  });

  function finishIfDone() {
    if (settled) return;
    if (fatalError && activeCollections === 0) {
      settled = true;
      doneReject(fatalError);
      return;
    }
    if (!fatalError && activeCollections === 0 && collectionQueue.length === 0) {
      settled = true;
      doneResolve();
    }
  }

  function enqueueCollection(collectionRef, depth) {
    const collectionPath = String(collectionRef?.path || "").trim();
    if (!collectionPath || seenCollections.has(collectionPath)) return;
    seenCollections.add(collectionPath);
    collectionQueue.push({ collectionRef, depth });
    pumpCollections();
  }

  async function processCollection({ collectionRef, depth }) {
    collectionCount += 1;
    maxDepth = Math.max(maxDepth, depth);
    let lastSnapshot = null;

    while (!fatalError) {
      let query = collectionRef.orderBy(FieldPath.documentId()).limit(PAGE_SIZE);
      if (lastSnapshot) query = query.startAfter(lastSnapshot);
      const page = await query.get();
      if (page.empty) break;

      const records = [];
      for (const snapshot of page.docs) {
        if (seenDocuments.has(snapshot.ref.path)) continue;
        seenDocuments.add(snapshot.ref.path);
        documentCount += 1;
        records.push({
          path: snapshot.ref.path,
          data: encode(snapshot.data()),
          createTime: timestampIso(snapshot.createTime),
          updateTime: timestampIso(snapshot.updateTime),
          readTime: timestampIso(snapshot.readTime)
        });
        if (documentCount % 500 === 0) {
          console.log(`[Firestore backup] ${documentCount} documenti letti, ${collectionCount} collezioni scoperte...`);
        }
      }
      await Promise.all(records.map(writeRecord));

      const childGroups = await Promise.all(page.docs.map((snapshot) =>
        childListLimit(() => snapshot.ref.listCollections())
      ));
      for (const children of childGroups) {
        for (const child of children) enqueueCollection(child, depth + 1);
      }

      lastSnapshot = page.docs[page.docs.length - 1];
      if (page.size < PAGE_SIZE) break;
    }
  }

  function pumpCollections() {
    if (settled) return;
    while (!fatalError && activeCollections < MAX_COLLECTION_WORKERS && collectionQueue.length) {
      const task = collectionQueue.shift();
      activeCollections += 1;
      processCollection(task)
        .catch((error) => {
          fatalError = fatalError || error;
        })
        .finally(() => {
          activeCollections -= 1;
          pumpCollections();
          finishIfDone();
        });
    }
    finishIfDone();
  }

  for (const collectionRef of topLevelCollections) enqueueCollection(collectionRef, 1);
  if (!topLevelCollections.length) finishIfDone();
  await allCollectionsDone;
  await writeChain;

  await new Promise((resolve, reject) => {
    out.once("error", reject);
    out.end(resolve);
  });

  const stat = fs.statSync(jsonlPath);
  const completedAt = new Date().toISOString();
  const manifest = {
    schemaVersion: 2,
    format: "VARGA_FIRESTORE_LOGICAL_JSONL_V1",
    projectId,
    databaseId: "(default)",
    startedAt,
    completedAt,
    transactionalSnapshot: false,
    note: "Snapshot logico completo: ogni documento conserva tipi Firestore e percorso. Il range startedAt/completedAt documenta l'intervallo di lettura.",
    documentCount,
    collectionCount,
    topLevelCollectionCount: topLevelCollections.length,
    maxDepth,
    bytes: stat.size,
    dataFile: "documents.jsonl",
    traversal: {
      collectionWorkers: MAX_COLLECTION_WORKERS,
      childListWorkers: MAX_CHILD_LIST_CALLS,
      pageSize: PAGE_SIZE
    }
  };
  fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);

  console.log(`Firestore logical backup completed: ${documentCount} documents, ${collectionCount} collections, ${stat.size} bytes.`);
  console.log(`Manifest: ${manifestPath}`);
}

main().then(() => process.exit(0)).catch((error) => {
  console.error("Firestore logical backup failed:", error?.stack || error?.message || error);
  process.exit(1);
});

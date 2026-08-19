#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");

const admin = require(path.join(process.cwd(), "functions", "node_modules", "firebase-admin"));

const projectId = String(process.env.PROJECT_ID || "hera-app-6cd2b").trim();
const outputDir = path.resolve(process.argv[2] || path.join(process.cwd(), "firestore-logical-backup"));
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

async function main() {
  const startedAt = new Date().toISOString();
  const jsonlPath = path.join(outputDir, "documents.jsonl");
  const manifestPath = path.join(outputDir, "manifest.json");
  const out = fs.createWriteStream(jsonlPath, { encoding: "utf8", flags: "w" });
  const seen = new Set();
  const topLevelCollections = await db.listCollections();
  let documentCount = 0;
  let collectionCount = 0;
  let maxDepth = 0;

  async function writeRecord(record) {
    if (!out.write(`${JSON.stringify(record)}\n`)) {
      await new Promise((resolve) => out.once("drain", resolve));
    }
  }

  async function walkCollection(collectionRef, depth) {
    collectionCount += 1;
    maxDepth = Math.max(maxDepth, depth);
    let lastSnapshot = null;
    const pageSize = 250;

    while (true) {
      let query = collectionRef.orderBy(FieldPath.documentId()).limit(pageSize);
      if (lastSnapshot) query = query.startAfter(lastSnapshot);
      const page = await query.get();
      if (page.empty) break;

      for (const snapshot of page.docs) {
        if (seen.has(snapshot.ref.path)) continue;
        seen.add(snapshot.ref.path);
        documentCount += 1;
        await writeRecord({
          path: snapshot.ref.path,
          data: encode(snapshot.data()),
          createTime: timestampIso(snapshot.createTime),
          updateTime: timestampIso(snapshot.updateTime),
          readTime: timestampIso(snapshot.readTime)
        });

        const children = await snapshot.ref.listCollections();
        for (const child of children) await walkCollection(child, depth + 1);
      }

      lastSnapshot = page.docs[page.docs.length - 1];
      if (page.size < pageSize) break;
    }
  }

  for (const collectionRef of topLevelCollections) await walkCollection(collectionRef, 1);

  await new Promise((resolve, reject) => {
    out.end(resolve);
    out.on("error", reject);
  });

  const stat = fs.statSync(jsonlPath);
  const completedAt = new Date().toISOString();
  const manifest = {
    schemaVersion: 1,
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
    dataFile: "documents.jsonl"
  };
  fs.writeFileSync(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`);

  console.log(`Firestore logical backup completed: ${documentCount} documents, ${collectionCount} collections, ${stat.size} bytes.`);
  console.log(`Manifest: ${manifestPath}`);
}

main().then(() => process.exit(0)).catch((error) => {
  console.error("Firestore logical backup failed:", error?.message || error);
  process.exit(1);
});

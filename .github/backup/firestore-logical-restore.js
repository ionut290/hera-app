#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const readline = require("node:readline");
const { createRequire } = require("node:module");

const requireFromFunctions = createRequire(path.join(process.cwd(), "functions", "package.json"));
const admin = requireFromFunctions("firebase-admin");
const { getFirestore, Timestamp, GeoPoint } = requireFromFunctions("firebase-admin/firestore");

const LABEL = "VARGA-STABLE-2026-08-19";
const projectId = String(process.env.PROJECT_ID || "hera-app-6cd2b").trim();
const inputDir = path.resolve(process.argv[2] || process.env.BACKUP_DIR || "firestore-logical-backup");
const manifestPath = path.join(inputDir, "manifest.json");
const dataPath = path.join(inputDir, "documents.jsonl");
const dryRun = String(process.env.DRY_RUN ?? "true").toLowerCase() !== "false";
const targetDatabaseId = String(process.env.TARGET_DATABASE_ID || "hera-restore-20260819").trim();
const exactRestore = String(process.env.EXACT_RESTORE ?? "true").toLowerCase() !== "false";
const confirmRestore = String(process.env.CONFIRM_RESTORE || "").trim();
const allowDefault = String(process.env.ALLOW_DEFAULT_DATABASE || "").trim();

if (!fs.existsSync(manifestPath) || !fs.existsSync(dataPath)) {
  throw new Error(`Backup incompleto: servono ${manifestPath} e ${dataPath}`);
}

const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));
if (manifest.format !== "VARGA_FIRESTORE_LOGICAL_JSONL_V1") {
  throw new Error(`Formato backup non riconosciuto: ${manifest.format || "mancante"}`);
}
if (manifest.projectId && manifest.projectId !== projectId) {
  throw new Error(`Il backup appartiene a ${manifest.projectId}, non a ${projectId}.`);
}
if (!admin.apps.length) {
  admin.initializeApp({
    credential: admin.credential.applicationDefault(),
    projectId
  });
}

const db = targetDatabaseId === "(default)" ? getFirestore() : getFirestore(targetDatabaseId);
try { db.settings({ ignoreUndefinedProperties: true }); } catch (_) {}

function decode(node) {
  if (!node || typeof node !== "object" || !node.t) {
    throw new Error("Valore backup non tipizzato o corrotto.");
  }
  switch (node.t) {
    case "null": return null;
    case "undefined": return undefined;
    case "string": return String(node.v ?? "");
    case "boolean": return Boolean(node.v);
    case "number": return Number(node.v);
    case "number-special":
      if (node.v === "NaN") return Number.NaN;
      if (node.v === "Infinity") return Infinity;
      if (node.v === "-Infinity") return -Infinity;
      throw new Error(`Numero speciale non riconosciuto: ${node.v}`);
    case "bigint": {
      const asBigInt = BigInt(String(node.v));
      const asNumber = Number(asBigInt);
      if (!Number.isSafeInteger(asNumber)) {
        throw new Error(`BigInt fuori dal range sicuro JavaScript: ${node.v}`);
      }
      return asNumber;
    }
    case "timestamp": return new Timestamp(Number(node.s), Number(node.n || 0));
    case "geopoint": return new GeoPoint(Number(node.lat), Number(node.lng));
    case "reference": {
      const referencePath = String(node.path || "").trim();
      if (!referencePath) throw new Error("DocumentReference senza path.");
      return db.doc(referencePath);
    }
    case "bytes": return Buffer.from(String(node.b64 || ""), "base64");
    case "date": return new Date(String(node.iso));
    case "array": return (Array.isArray(node.v) ? node.v : []).map((entry) => {
      const value = decode(entry);
      return value === undefined ? null : value;
    });
    case "map": {
      const result = {};
      for (const [key, entry] of Object.entries(node.v || {})) {
        const value = decode(entry);
        if (value !== undefined) result[key] = value;
      }
      return result;
    }
    default: throw new Error(`Tipo backup non riconosciuto: ${node.t}`);
  }
}

function validateDocumentPath(documentPath) {
  const clean = String(documentPath || "").trim();
  const parts = clean.split("/").filter(Boolean);
  if (!clean || parts.length < 2 || parts.length % 2 !== 0) {
    throw new Error(`Path documento Firestore non valido: ${documentPath}`);
  }
  return clean;
}

async function scanBackup(onRecord) {
  const input = fs.createReadStream(dataPath, { encoding: "utf8" });
  const lines = readline.createInterface({ input, crlfDelay: Infinity });
  let count = 0;
  let decodedBytes = 0;
  const topCollections = new Set();

  for await (const line of lines) {
    if (!line.trim()) continue;
    const record = JSON.parse(line);
    record.path = validateDocumentPath(record.path);
    record.decoded = decode(record.data);
    topCollections.add(record.path.split("/")[0]);
    count += 1;
    decodedBytes += Buffer.byteLength(line, "utf8");
    if (onRecord) await onRecord(record, count);
  }
  return { count, decodedBytes, topCollections: [...topCollections].sort() };
}

async function clearTargetDatabase() {
  const collections = await db.listCollections();
  if (!collections.length) return { collectionsDeleted: 0 };
  if (typeof db.recursiveDelete !== "function") {
    throw new Error("Questa versione Firestore non espone recursiveDelete: cancellazione esatta interrotta per sicurezza.");
  }
  let deleted = 0;
  for (const collection of collections) {
    await db.recursiveDelete(collection);
    deleted += 1;
  }
  return { collectionsDeleted: deleted };
}

async function restore() {
  console.log(`Recovery label: ${LABEL}`);
  console.log(`Project: ${projectId}`);
  console.log(`Target database: ${targetDatabaseId}`);
  console.log(`Mode: ${dryRun ? "DRY RUN" : "WRITE"}`);
  console.log(`Exact restore: ${exactRestore}`);

  if (!dryRun) {
    if (confirmRestore !== LABEL) {
      throw new Error(`Scrittura bloccata: impostare CONFIRM_RESTORE=${LABEL}`);
    }
    if (targetDatabaseId === "(default)" && allowDefault !== "DELETE_AND_RESTORE_VARGA_2026_08_19") {
      throw new Error("Database principale protetto: manca ALLOW_DEFAULT_DATABASE=DELETE_AND_RESTORE_VARGA_2026_08_19");
    }
  }

  const preflight = await scanBackup();
  if (Number(manifest.documentCount || 0) && preflight.count !== Number(manifest.documentCount)) {
    throw new Error(`Conteggio backup non coerente: manifest=${manifest.documentCount}, JSONL=${preflight.count}`);
  }
  console.log(`Preflight OK: ${preflight.count} documenti, ${preflight.topCollections.length} collezioni top-level.`);

  if (dryRun) {
    console.log("DRY RUN completato: nessuna scrittura e nessuna cancellazione eseguita.");
    return { mode: "dry-run", ...preflight };
  }

  let deleteResult = { collectionsDeleted: 0 };
  if (exactRestore) {
    console.log("Cancellazione controllata del database destinazione per ripristino esatto...");
    deleteResult = await clearTargetDatabase();
  }

  const writer = db.bulkWriter();
  writer.onWriteError((error) => {
    if (error.failedAttempts < 5) return true;
    console.error(`Scrittura fallita definitivamente: ${error.documentRef?.path || "documento"}`, error.message);
    return false;
  });

  let scheduled = 0;
  let written = 0;
  writer.onWriteResult(() => { written += 1; });

  try {
    await scanBackup(async (record) => {
      writer.set(db.doc(record.path), record.decoded);
      scheduled += 1;
      if (scheduled % 1000 === 0) console.log(`Ripristino programmato: ${scheduled}/${preflight.count}`);
    });
    await writer.close();
  } catch (error) {
    try { await writer.close(); } catch (_) {}
    throw error;
  }

  if (written !== preflight.count) {
    throw new Error(`Ripristino incompleto: scritti ${written}/${preflight.count} documenti.`);
  }

  const result = {
    mode: "write",
    exactRestore,
    targetDatabaseId,
    expectedDocuments: preflight.count,
    writtenDocuments: written,
    deletedTopLevelCollections: deleteResult.collectionsDeleted,
    completedAt: new Date().toISOString()
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

restore().then(() => process.exit(0)).catch((error) => {
  console.error("Ripristino Firestore interrotto:", error?.stack || error?.message || error);
  process.exit(1);
});

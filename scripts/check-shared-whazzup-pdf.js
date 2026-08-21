"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const moduleCode = fs.readFileSync("shared-pdf-attachments.js", "utf8");
const loaderCode = fs.readFileSync("loading-humor.js", "utf8");
const cleanupCode = fs.readFileSync("functions/cleanup-whazzup-pdfs.js", "utf8");
const mainCode = fs.readFileSync("functions/main.js", "utf8");

[
  'pdfButton.dataset.photoSource = "pdf"',
  "Aggiungi PDF",
  "Documento condiviso con tutti",
  'const SOURCE = "whazzup-impianto-pdf";',
  'visibility: "global"',
  "expiresAt:",
  "autoDeleteAfterDays: 30",
  'storage.ref().child(storagePath)',
  '.where("impiantoKey", "==", impiantoKey)',
  "PDF condivisi"
].forEach((marker) => assert.ok(moduleCode.includes(marker), `Marker PDF condiviso mancante: ${marker}`));

assert.ok(loaderCode.includes("shared-pdf-attachments.js"), "Loader del modulo PDF condiviso mancante");
assert.ok(cleanupCode.includes("cleanupExpiredWhazzupPdfs"), "Cleanup schedulato PDF mancante");
assert.ok(cleanupCode.includes('.where("source", "==", SOURCE)'), "Cleanup non limita la lettura ai PDF Whazzup");
assert.ok(cleanupCode.includes("expiresAtMs > now"), "Cleanup non verifica la scadenza a 30 giorni");
assert.ok(mainCode.includes('require("./cleanup-whazzup-pdfs")'), "Funzione cleanup non esportata dal main Functions");

console.log("Shared Whazzup PDF checks passed.");

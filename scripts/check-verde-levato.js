#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("verde-levato.js", "utf8");
const loader = fs.readFileSync("app-pure-utils.js", "utf8");
const html = fs.readFileSync("index.html", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");
const rules = fs.readFileSync("firestore.rules", "utf8");

assert.match(source, /title: "Cantieri"/);
assert.match(source, /title: "Alberi censiti"/);
assert.match(source, /title: "Siepi"/);
assert.match(source, /RECORDS_COLLECTION = "verdeLevatoRecords"/);
assert.match(source, /COMMESSE_COLLECTION = "verdeLevatoCommesse"/);
assert.match(source, /CONFIG_COLLECTION = "verdeLevatoConfig"/);
assert.match(source, /source: "MANUALE_VERDE_LEVATO"/);
assert.match(loader, /addScript\("\.\/verde-levato\.js\?v=20260902-verde-levato2", "hera-verde-levato"\)/);
assert.match(source, /LA MIA POSIZIONE/);
assert.match(source, /enableHighAccuracy: true, timeout: 18000, maximumAge: 0/);
assert.match(source, /nominatim\.openstreetmap\.org\/reverse/);
for (const field of ["comune", "localita", "via", "civico", "cap", "provincia", "regione", "paese", "indirizzo"]) {
  assert.match(source, new RegExp(`${field}:`), `campo automatico ${field} mancante`);
}
assert.match(source, /specieAlbero/);
assert.match(source, /specieSiepe/);
assert.match(source, /lavorazioniRichieste/);
assert.match(source, /name="commessaId" required/);
assert.match(source, /NUOVA COMMESSA/);
assert.match(source, /commessaId: selectedCommessa\?\.id \|\| ""/);
assert.match(source, /commessaNome: text\(selectedCommessa\?\.nome\)/);
assert.match(source, /Seleziona o crea la commessa Verde Levato da associare al cantiere/);
assert.match(source, /ESPORTA TUTTI I DATI/);
assert.match(source, /HeraHeavyLibs\?\.ensure\?\.\("xlsx"\)/);
assert.match(source, /verde_levato_dati_completi_/);
assert.match(source, /book_append_sheet\(workbook, dataSheet, "Dati completi"\)/);
assert.match(source, /book_append_sheet\(workbook, commesseSheet, "Commesse"\)/);
assert.match(source, /state\.globalAdmin \|\| \(email && state\.adminEmails\.includes\(email\)\)/);
assert.match(source, /Non diventerà amministratore generale dell’app/);
assert.doesNotMatch(source, /onSnapshot\s*\(/, "Verde Levato non deve creare listener Firestore persistenti");
assert.doesNotMatch(source, /\.delete\s*\(/, "Verde Levato non deve cancellare dati");
assert.doesNotMatch(source, /markImpiantoDone|handleImpiantoWhatsAppClick|setImpiantoDone/);

assert.match(rules, /function isVerdeLevatoAdmin\(\)/);
assert.match(rules, /match \/verdeLevatoConfig\/\{documentId\}/);
assert.match(rules, /allow create, update: if isAdmin\(\)/);
assert.match(rules, /match \/verdeLevatoRecords\/\{recordId\}/);
assert.match(rules, /allow create, update: if isVerdeLevatoAdmin\(\)/);
assert.match(rules, /allow delete: if false/);
assert.match(rules, /match \/verdeLevatoCommesse\/\{commessaId\}/);

assert.match(loader, /verde-levato\.js\?v=20260902-verde-levato2/);
assert.match(html, /app-pure-utils\.js\?v=20260902-verde-levato2/);
assert.match(sw, /verde-levato\.js\?v=20260902-verde-levato2/);
assert.match(sw, /"\/verde-levato\.js"/);

console.log("✅ Verde Levato: commesse nel modulo, associazione cantieri, Excel, GPS e sicurezza verificati.");

#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const rules = fs.readFileSync("firestore.rules", "utf8");
const fattoSource = fs.readFileSync("fatto-button-immediate.js", "utf8");

assert.match(rules, /^rules_version = '2';/m, "Le regole Firestore devono restare in versione 2");

function sliceBetween(startMarker, endMarker) {
  const start = rules.indexOf(startMarker);
  assert.notEqual(start, -1, `Blocco non trovato: ${startMarker}`);
  const end = rules.indexOf(endMarker, start + startMarker.length);
  assert.notEqual(end, -1, `Fine blocco non trovata: ${endMarker}`);
  return rules.slice(start, end);
}

const commesseBlock = sliceBetween(
  "match /commesse/{commessaId}",
  "// Il report contiene dati storici"
);

assert.match(
  commesseBlock,
  /allow read:\s*if signedIn\(\);[\s\S]*allow write:\s*if isAdmin\(\);/,
  "La commessa padre deve essere leggibile dagli utenti ma scrivibile solo dagli admin"
);
assert.doesNotMatch(
  commesseBlock,
  /allow write:\s*if signedIn\(\);/,
  "Un utente normale non deve poter modificare la commessa padre"
);
assert.match(
  commesseBlock,
  /match \/\{subcollection\}\/\{document=\*\*\}/,
  "Le sotto-collezioni operative devono avere un match che esclude il documento padre"
);
assert.doesNotMatch(
  commesseBlock,
  /match \/\{document=\*\*\}/,
  "Con rules v2 il wildcard ricorsivo non deve partire direttamente dalla commessa padre"
);
assert.match(
  commesseBlock,
  /match \/\{subcollection\}\/\{document=\*\*\}[\s\S]*allow read, write:\s*if signedIn\(\);/,
  "FATTO e lavorazioni devono continuare a poter scrivere nei documenti figli"
);

const globalBlock = sliceBetween(
  "match /globalCommesse/{commessaId}",
  "match /{collection}/{docId}"
);
assert.match(
  globalBlock,
  /allow read:\s*if signedIn\(\);[\s\S]*allow write:\s*if isAdmin\(\);/,
  "La commessa Global padre deve essere scrivibile solo dagli admin"
);
assert.match(
  globalBlock,
  /match \/\{subcollection\}\/\{document=\*\*\}[\s\S]*allow read, write:\s*if signedIn\(\);/,
  "I dati figli Global devono restare operativi per gli utenti autenticati"
);

const genericCatchAll = rules.slice(rules.indexOf("match /{collection}/{docId}"));
assert.doesNotMatch(
  genericCatchAll,
  /"globalCommesse"/,
  "globalCommesse non deve essere riaperta dal catch-all generico"
);

for (const collectionName of ["lavorazioni", "impiantiFisici", "impianti"]) {
  assert.match(
    fattoSource,
    new RegExp(`collection\\(\\"${collectionName}\\"\\)`),
    `Il flusso FATTO deve continuare a usare la sotto-collezione ${collectionName}`
  );
}

console.log("✅ Sicurezza Firestore commesse: padri admin-only, sotto-collezioni operative preservate.");

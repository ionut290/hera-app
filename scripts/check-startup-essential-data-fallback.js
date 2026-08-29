#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const optimizer = fs.readFileSync("firestore-startup-cost-optimizer.js", "utf8");
const firebaseConfig = fs.readFileSync("firebase-config.js", "utf8");

assert.match(
  app,
  /function subscribeCommesse\(\)[\s\S]*?runFirestoreGetWithRetry\(query,[\s\S]*?loadCommesseFromLocalCache\(\)/,
  "Le commesse devono usare Firestore e mantenere il recupero locale in caso di errore."
);
assert.match(
  app,
  /if \(!incrementalState \|\| !cachedImpianti\.length\) \{\s*unsubscribeImpianti = startFullListener\(\)/,
  "Gli impianti devono aprire il listener Firestore completo quando la cache manca."
);
assert.match(
  optimizer,
  /if \(!Array\.isArray\(primaryRows\)\) return;/,
  "Una vista squadre vuota deve essere considerata un risultato valido."
);
assert.doesNotMatch(
  optimizer,
  /vista-principale-vuota/,
  "Una giornata senza squadre non deve causare una lunga ricerca Firestore."
);
assert.match(
  optimizer,
  /selected, today, tomorrow, \.\.\.nextWorkdayCandidates/,
  "La cache squadre deve preparare oggi, domani e il prossimo giorno lavorativo."
);
assert.match(
  optimizer,
  /const SQUADRE_FALLBACK_MS = 1200/,
  "Se la vista manca davvero, il recupero Firestore deve partire rapidamente."
);
assert.match(
  optimizer,
  /unsubs\.splice\(0\)\.forEach[\s\S]*?nativeUnsubscribe = originalOnSnapshot\.apply\(query, args\)/,
  "Il fallback squadre deve chiudere le viste condivise prima di aprire il listener Firestore."
);
assert.match(
  firebaseConfig,
  /firestore-startup-cost-optimizer\.js\?v=20260829-fast-squad-cache1/,
  "Il browser deve ricevere la nuova versione del fallback dati essenziali."
);

console.log("Dati essenziali post-login: fallback Firestore per commesse, squadre e impianti verificato.");

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
  /if \(!Array\.isArray\(primaryRows\) \|\| primaryRows\.length === 0\) \{\s*startFallback\("vista-principale-vuota"\)/,
  "Una vista squadre vuota deve attivare il fallback Firestore."
);
assert.match(
  optimizer,
  /unsubs\.splice\(0\)\.forEach[\s\S]*?nativeUnsubscribe = originalOnSnapshot\.apply\(query, args\)/,
  "Il fallback squadre deve chiudere le viste condivise prima di aprire il listener Firestore."
);
assert.match(
  firebaseConfig,
  /firestore-startup-cost-optimizer\.js\?v=20260828-device-cache-fallback1/,
  "Il browser deve ricevere la nuova versione del fallback dati essenziali."
);

console.log("Dati essenziali post-login: fallback Firestore per commesse, squadre e impianti verificato.");

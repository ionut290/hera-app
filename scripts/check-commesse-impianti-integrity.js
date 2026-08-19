#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync("operational-import-repair.js", "utf8");
const timers = [];
const sandbox = {
  window: {
    addEventListener() {},
    removeEventListener() {}
  },
  document: {
    hidden: false,
    addEventListener() {},
    removeEventListener() {}
  },
  console,
  setTimeout(callback, delay) {
    timers.push({ callback, delay, type: "timeout" });
    return timers.length;
  },
  setInterval(callback, delay) {
    timers.push({ callback, delay, type: "interval" });
    return timers.length;
  },
  clearTimeout() {},
  clearInterval() {}
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox, { filename: "operational-import-repair.js" });

const api = sandbox.window.INRETE_MODENA_AUGUST_2026;
assert.ok(api, "Runtime INRETE Modena non esportato");
assert.equal(api.plants.length, 30, "Il dataset deve contenere esattamente 30 impianti");
assert.equal(new Set(api.plants.map((plant) => plant.id)).size, 30, "Gli ID impianto devono essere univoci");

for (const plant of api.plants) {
  const pair = api.testing.coordinatePairFrom(plant);
  assert.ok(pair, `Coordinate mancanti per ${plant.id}`);
  assert.ok(pair.lat >= 43.5 && pair.lat <= 45.2, `Latitudine fuori area per ${plant.id}`);
  assert.ok(pair.lng >= 9.5 && pair.lng <= 12.5, `Longitudine fuori area per ${plant.id}`);
}

assert.deepEqual(
  Array.from(api.testing.detectKinds({ tipologiaLavorazione: "STRAORDINARIO" })),
  ["STRAORDINARIO"],
  "STRAORDINARIO non deve essere scambiato per ORDINARIO"
);
assert.deepEqual(
  Array.from(api.testing.detectKinds({ tipologiaLavorazione: "MANUTENZIONE ORDINARIA" })),
  ["ORDINARIO"]
);
assert.deepEqual(
  new Set(api.testing.detectKinds({ tipologiaLavorazione: "ORDINARIO E STRAORDINARIO" })),
  new Set(["ORDINARIO", "STRAORDINARIO"])
);

const canonical = api.plants.find((plant) => plant.id === "sap_3430707");
assert.ok(canonical);

api.testing.setWorkItems([]);
let merged = api.mergePlants(api.plants, [{
  id: canonical.id,
  idSap: canonical.idSap,
  denominazione: canonical.denominazione,
  comune: canonical.comune,
  indirizzo: canonical.indirizzo,
  latitudine: 44.60001,
  longitudine: 11.05001,
  stato: "FATTO",
  statoGenerale: "FATTO",
  done: true
}]);
let target = merged.find((plant) => plant.idSap === canonical.idSap);
assert.equal(target.statoGenerale, "FATTO", "Lo stato remoto deve prevalere sul fallback DA FARE");
assert.equal(target.done, true, "Il flag done remoto non deve essere azzerato");
assert.equal(target.latitudine, 44.60001, "Le coordinate remote valide devono prevalere sul fallback");
assert.equal(target.longitudine, 11.05001);

const fallbackSnapshot = api.plants.map((plant) => ({ ...plant }));
merged = api.mergePlants(api.plants, [{
  ...canonical,
  localFallback: false,
  latitudine: 44.60001,
  longitudine: 11.05001,
  gpsY: 44.60001,
  gpsX: 11.05001,
  stato: "FATTO",
  statoGenerale: "FATTO",
  done: true,
  updatedAt: { seconds: 1, nanoseconds: 0 }
}], fallbackSnapshot);
target = merged.find((plant) => plant.id === canonical.id);
assert.equal(target.statoGenerale, "FATTO", "Il fallback corrente non deve prevalere sul dato Firestore");
assert.equal(target.done, true);
assert.equal(target.latitudine, 44.60001);
assert.equal(target.longitudine, 11.05001);
assert.equal(target.localFallback, false);

api.testing.setWorkItems([{
  id: "work-extra",
  impiantoId: canonical.id,
  tipologiaLavorazione: "STRAORDINARIO",
  stato: "DA FARE"
}]);
merged = api.mergePlants(api.plants);
target = merged.find((plant) => plant.id === canonical.id);
assert.equal(target.attivitaLabel, "STRAORDINARIO", "Una sola lavorazione straordinaria non deve risultare mista");
assert.equal(target.isMixedOrdinaryExtraordinary, false);

api.testing.setWorkItems([
  {
    id: "work-ordinary",
    impiantoId: canonical.id,
    tipologiaLavorazione: "ORDINARIO",
    stato: "FATTO"
  },
  {
    id: "work-extra",
    impiantoId: canonical.id,
    tipologiaLavorazione: "STRAORDINARIO",
    stato: "DA FARE"
  }
]);
merged = api.mergePlants(api.plants);
target = merged.find((plant) => plant.id === canonical.id);
assert.equal(target.attivitaLabel, "ORDINARIO E STRAORDINARIO");
assert.equal(target.markerClass, "impianto-marker-mixed-yellow");
assert.equal(target.statoGenerale, "PARZIALMENTE FATTO");
assert.equal(target.numeroLavorazioni, 2);
assert.equal(target.numeroLavorazioniFatte, 1);
assert.equal(target.numeroLavorazioniDaFare, 1);

const signatureBefore = api.testing.signature([canonical]);
const signatureAfter = api.testing.signature([{ ...canonical, latitudine: 44.601, gpsY: 44.601 }]);
assert.notEqual(signatureBefore, signatureAfter, "La firma UI deve cambiare quando cambiano le coordinate");

const safePatch = api.testing.buildExistingPlantPatch({
  id: canonical.id,
  stato: "FATTO",
  statoGenerale: "FATTO",
  done: true,
  latitudine: canonical.latitudine,
  longitudine: canonical.longitudine
}, canonical, api.testing.deriveWorkAggregate([]), canonical.id);
assert.equal(Object.prototype.hasOwnProperty.call(safePatch, "stato"), false, "Il backfill non deve riscrivere lo stato esistente");
assert.equal(Object.prototype.hasOwnProperty.call(safePatch, "statoGenerale"), false);
assert.equal(Object.prototype.hasOwnProperty.call(safePatch, "done"), false);
assert.equal(safePatch.localFallback, false, "Il documento Firestore non deve restare marcato come fallback locale");

const recoveryPatch = api.testing.buildExistingPlantPatch({
  id: canonical.id,
  stato: "DA FARE",
  statoGenerale: "DA FARE",
  done: false
}, canonical, api.testing.deriveWorkAggregate([
  { stato: "FATTO", tipologiaLavorazione: "ORDINARIO" },
  { stato: "FATTO", tipologiaLavorazione: "STRAORDINARIO" }
]), canonical.id);
assert.equal(recoveryPatch.statoGenerale, "FATTO", "Le lavorazioni completate devono recuperare uno stato impianto azzerato");
assert.equal(recoveryPatch.done, true);

const swapped = api.testing.coordinatePairFrom({ latitudine: 10.83482, longitudine: 44.53632 });
assert.deepEqual(
  { lat: swapped.lat, lng: swapped.lng },
  { lat: 44.53632, lng: 10.83482 },
  "Le coordinate X/Y invertite devono essere corrette"
);

const sameNamePlants = api.plants.filter((plant) => plant.denominazione === "REMI S.CLEMENTE");
assert.equal(sameNamePlants.length, 2);
api.testing.setWorkItems([]);
merged = api.mergePlants(api.plants);
assert.equal(
  merged.filter((plant) => plant.denominazione === "REMI S.CLEMENTE").length,
  2,
  "Due impianti con stessa anagrafica ma ID SAP diversi non devono essere accorpati"
);

assert.match(source, /buildExistingPlantPatch/);
assert.doesNotMatch(
  source,
  /batch\.set\(ref\.collection\("impianti"\)\.doc\(plant\.id\),\s*data,\s*\{\s*merge:\s*true\s*\}\)/,
  "Non deve tornare la riscrittura cieca di tutti gli impianti"
);
assert.match(source, /String\(existingParentData\.datasetVersion \|\| ""\) === DATASET_VERSION/);
assert.match(source, /document\.hidden/);

let historyStops = 0;
let historySubscriptions = 0;
const restoredOnSnapshot = function restoredOnSnapshot() {};
const Query = function Query() {};
Query.prototype.onSnapshot = restoredOnSnapshot;
Query.prototype.__heraActiveCommesseOriginalOnSnapshot = restoredOnSnapshot;
const historySandbox = {
  currentUser: { uid: "tester" },
  stopCommesseSubscription() {
    historyStops += 1;
  },
  subscribeCommesse() {
    historySubscriptions += 1;
    return Promise.resolve(true);
  },
  window: {
    firebase: {
      firestore: { Query },
      auth() {
        return {
          currentUser: { uid: "tester" },
          onAuthStateChanged() {
            return () => {};
          }
        };
      }
    },
    addEventListener() {}
  },
  document: {
    hidden: false,
    addEventListener() {}
  },
  console,
  setTimeout() { return 1; },
  setInterval() { return 1; },
  clearTimeout() {},
  clearInterval() {}
};
vm.createContext(historySandbox);
vm.runInContext(source, historySandbox, { filename: "operational-import-repair-history.js" });
assert.ok(historySandbox.window.HeraHistoricalCommesseResubscribe, "Manca ripristino listener commesse storiche");
assert.equal(historySandbox.window.HeraHistoricalCommesseResubscribe.refresh(), true);
assert.equal(historyStops, 1);
assert.equal(historySubscriptions, 1);
assert.equal(historySandbox.window.HeraHistoricalCommesseResubscribe.refresh(), true);
assert.equal(historyStops, 1, "Il listener storico non deve essere duplicato");
assert.equal(historySubscriptions, 1);

assert.match(source, /installHistoricalCommesseResubscribe/);
assert.match(source, /QueryPrototype\.onSnapshot === restored/);

const statsCacheSource = fs.readFileSync("commessa-stats-cache-optimizer.js", "utf8");
assert.match(statsCacheSource, /function activeCollectionName\(\)/, "La cache statistiche deve risolvere la raccolta commesse attiva");
assert.match(statsCacheSource, /function commessaRef\(commessaId\)/, "Manca il riferimento centralizzato alla commessa attiva");
assert.doesNotMatch(
  statsCacheSource,
  /db\.collection\("commesse"\)\.doc\(commessaId\)/,
  "La cache statistiche non deve leggere una raccolta commesse hard-coded"
);

const globalArchiveSource = fs.readFileSync("global-archive-sync.js", "utf8");
assert.match(globalArchiveSource, /async function ensureVisibleCommessa\(commessaId\)/, "Manca la protezione anagrafica della commessa Global");
const mirrorVisiblePlantBody = globalArchiveSource.match(/async function mirrorVisiblePlant\(commessaId, plantId, raw = \{\}\) \{[\s\S]*?\n  \}/)?.[0] || "";
assert.ok(mirrorVisiblePlantBody, "Funzione mirrorVisiblePlant non trovata");
assert.match(mirrorVisiblePlantBody, /await ensureVisibleCommessa\(commessaId\)/, "L'impianto deve garantire l'esistenza della commessa senza usarne i propri dati");
assert.doesNotMatch(
  mirrorVisiblePlantBody,
  /mirrorVisibleCommessa\(commessaId,\s*raw\)/,
  "I dati dell'impianto non devono mai aggiornare l'anagrafica della commessa Global"
);

console.log("✅ Audit commesse-impianti: merge, coordinate, stati, lavorazioni, cache, Global e backfill non distruttivo verificati.");

#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync("operational-import-repair.js", "utf8");
let renderedHome = 0;
let renderedManagement = 0;
const maps = {
  commesseById: new Map([
    ["inrete_modena_agosto_2026", { id: "inrete_modena_agosto_2026", nome: "INRETE MODENA - AGOSTO 2026", codice: "INRETE-MO-AGO-2026", attiva: true }],
    ["canonical-modena", { id: "canonical-modena", nome: "INRETE MODENA", codice: "28015", attiva: true, impiantiCount: 30, impiantiFattiCount: 2 }]
  ]),
  impiantiByCommessaId: new Map([["inrete_modena_agosto_2026", [{ id: "old" }]]]),
  commessaStatsById: new Map(),
  commessaWorkSummariesById: new Map(),
  commessaHoursById: new Map(),
  unsubscribeCommessaStats: new Map()
};

const sandbox = {
  console: { info() {}, warn() {}, error() {} },
  ...maps,
  currentUser: { uid: "u1" },
  selectedCommessaId: "",
  managementCommessaId: "",
  currentImpianti: [],
  getCommesseCollectionName: () => "commesse",
  canManageData: () => false,
  calculateImpiantiStats(items) { return { total: items.length, done: items.filter((x) => x.done).length }; },
  recalculateCommessaWorkSummaries() {},
  renderCommesseHomeList() { renderedHome += 1; },
  renderCommesseManagementList() { renderedManagement += 1; },
  renderParentCommessaOverview() {},
  updateCommessaDashboard() {},
  renderImpianti() {},
  renderMap() {},
  setTimeout() { return 1; },
  setInterval() { return 1; },
  clearTimeout() {},
  clearInterval() {},
  document: {
    hidden: false,
    addEventListener() {},
    querySelector() { return null; }
  },
  window: {
    addEventListener() {},
    firebase: { firestore: { Query: function Query() {} } }
  }
};
sandbox.window.window = sandbox.window;
vm.createContext(sandbox);
vm.runInContext(source, sandbox, { filename: "operational-import-repair.js" });

const api = sandbox.window.INRETE_MODENA_AUGUST_2026;
assert.ok(api, "Runtime INRETE Modena non esportato");
assert.equal(api.testing.findCanonicalModenaCommessa().id, "canonical-modena");
assert.equal(api.ensureVisibleLocally({ refresh: false }), true);
assert.equal(sandbox.commesseById.has("inrete_modena_agosto_2026"), false, "La commessa tecnica non deve restare nella UI");
assert.equal(sandbox.impiantiByCommessaId.has("inrete_modena_agosto_2026"), false, "La cache della commessa tecnica deve essere eliminata");
assert.equal(sandbox.commesseById.has("canonical-modena"), true, "La commessa 28015 deve restare disponibile");
assert.equal(sandbox.commessaStatsById.get("canonical-modena").total, 30, "La home deve usare subito il conteggio parent invece di 0");
assert.ok(renderedHome > 0 && renderedManagement > 0, "La UI deve essere ridisegnata dopo la pulizia");

const operationalPlants = Array.from({ length: 30 }, (_, index) => ({
  id: `plant-${index + 1}`,
  idSap: String(3430000 + index),
  denominazione: `Impianto ${index + 1}`,
  stato: "DA FARE",
  done: false
}));
const work = Array.from({ length: 42 }, (_, index) => ({
  id: `work-${index + 1}`,
  impiantoId: `plant-${(index % 30) + 1}`,
  idSap: String(3430000 + (index % 30)),
  stato: index < 4 ? "FATTO" : "DA FARE",
  totale: index < 4 ? 79.2 : 0
}));
const summary = api.testing.summarizeData(operationalPlants, [], work);
assert.equal(summary.uniquePlants, 30, "42 lavorazioni non devono diventare 42 impianti");
assert.equal(summary.workRows, 42);
assert.equal(summary.doneWork, 4);
assert.equal(summary.pendingWork, 38);
assert.equal(summary.items.length, 30);

const sameSap = api.testing.summarizeData([
  { id: "a", idSap: "3430707", denominazione: "REMI CASTELFRANCO" },
  { id: "b", idSap: "3430707", denominazione: "REMI CASTELFRANCO" }
], [], []);
assert.equal(sameSap.uniquePlants, 1, "Lo stesso ID SAP deve contare come un solo impianto fisico");

const sameNameDifferentSap = api.testing.summarizeData([
  { id: "a", idSap: "3425784", denominazione: "REMI S.CLEMENTE" },
  { id: "b", idSap: "3426714", denominazione: "REMI S.CLEMENTE" }
], [], []);
assert.equal(sameNameDifferentSap.uniquePlants, 2, "ID SAP diversi non devono essere accorpati anche se il nome coincide");

// Regressione: un riepilogo INRETE in cache non deve mai sovrascrivere un
// FATTO/RESET appena applicato dal listener operativo o dallo stato locale.
sandbox.selectedCommessaId = "canonical-modena";
sandbox.currentImpianti = [{ id: "plant-1", idSap: "3430707", done: true, stato: "FATTO" }];
sandbox.impiantiByCommessaId.set("canonical-modena", sandbox.currentImpianti.map((item) => ({ ...item })));
api.testing.applySummary(sandbox.commesseById.get("canonical-modena"), {
  uniquePlants: 1,
  donePlants: 0,
  todoPlants: 1,
  workRows: 1,
  doneWork: 0,
  pendingWork: 1,
  subtotalCompleted: 0,
  items: [{ id: "plant-1", idSap: "3430707", done: false, stato: "DA FARE" }]
});
assert.equal(sandbox.currentImpianti[0].done, true, "Un riepilogo vecchio non deve annullare FATTO nella lista aperta");
assert.equal(sandbox.impiantiByCommessaId.get("canonical-modena")[0].done, true, "Un riepilogo vecchio non deve avvelenare la cache operativa");

assert.match(source, /CANONICAL_COMMESSA_CODE = "28015"/);
assert.match(source, /archiveSyntheticDuplicate/);
assert.match(source, /hiddenFromHome: true/);
assert.match(source, /attiva: false/);
assert.doesNotMatch(source, /commesseById\.set\(SYNTHETIC_COMMESSA_ID/, "Il runtime non deve più ricreare la commessa tecnica");
assert.doesNotMatch(source, /ref\.delete\(/, "Il duplicato deve essere archiviato, non cancellato");
assert.match(source, /<b>\$\{summary\.uniquePlants\}<\/b> impianti/);
assert.match(source, /<b>\$\{summary\.workRows\}<\/b> lavorazioni/);
assert.doesNotMatch(source, /currentImpianti\s*=\s*summary\.items/, "Il riepilogo periodico non deve sostituire la lista operativa");
assert.doesNotMatch(source, /impiantiByCommessaId\.set\(canonical\.id,\s*summary\.items/, "Il riepilogo periodico non deve sostituire la cache operativa");
assert.doesNotMatch(source, /setInterval\s*\(/, "Il riepilogo INRETE non deve effettuare polling Firestore continuo");
assert.match(source, /duplicateArchiveChecked/, "La verifica del duplicato tecnico deve essere eseguita una sola volta per sessione");
assert.match(source, /MutationObserver/, "L'apertura della gestione deve essere rilevata tramite evento, non polling");

const statsCacheSource = fs.readFileSync("commessa-stats-cache-optimizer.js", "utf8");
assert.match(statsCacheSource, /function activeCollectionName\(\)/, "La cache statistiche deve usare la raccolta attiva");
assert.match(statsCacheSource, /docChanges\(\)/, "Il change index deve continuare a processare solo i delta");

const globalArchiveSource = fs.readFileSync("global-archive-sync.js", "utf8");
assert.match(globalArchiveSource, /async function ensureVisibleCommessa\(commessaId\)/, "La protezione Global deve restare attiva");

console.log("✅ INRETE Modena: una sola commessa 28015, conteggio impianti unici e lavorazioni separate verificati.");

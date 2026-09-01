#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync("operational-import-repair.js", "utf8");
const appSource = fs.readFileSync("app.js", "utf8");
const accountingSource = fs.readFileSync("accounting-v2.js", "utf8");

const commesseById = new Map([
  ["modena", { id: "modena", nome: "INRETE MODENA", codice: "28015", impiantiCount: 30 }],
  ["ferrara", { id: "ferrara", nome: "INRETE FERRARA", codice: "28015001", impiantiCount: 94 }],
  ["bologna", { id: "bologna", nome: "INRETE BOLOGNA", codice: "28015002", impiantiCount: 120 }]
]);

const meta = { textContent: "INRETE FERRARA • Cod. 28015001" };
const stats = { innerHTML: "FERRARA ORIGINALE" };
const managementScreen = { classList: { contains(name) { return name === "hidden" ? false : false; } } };

const sandbox = {
  console: { info() {}, warn() {}, error() {} },
  addEventListener() {},
  document: {
    hidden: false,
    addEventListener() {},
    querySelector(selector) {
      if (selector === "#impianti-management-meta") return meta;
      if (selector === "#impianti-management-stats") return stats;
      if (selector === "#impianti-management-screen") return managementScreen;
      return null;
    }
  },
  commesseById,
  impiantiByCommessaId: new Map(),
  commessaStatsById: new Map(),
  commessaWorkSummariesById: new Map(),
  commessaHoursById: new Map(),
  unsubscribeCommessaStats: new Map(),
  managementCommessaId: "ferrara",
  selectedCommessaId: "ferrara",
  currentImpianti: [],
  currentUser: { uid: "u1" },
  auth: { currentUser: { uid: "u1" }, onAuthStateChanged() {} },
  getCommesseCollectionName() { return "commesse"; },
  canManageData() { return false; },
  renderCommesseHomeList() {},
  renderCommesseManagementList() {},
  recalculateCommessaWorkSummaries() {},
  renderParentCommessaOverview() {},
  updateCommessaDashboard() {},
  renderImpianti() {},
  renderMap() {},
  stopCommesseSubscription() {},
  subscribeCommesse() {},
  firebase: { firestore: { Query: function Query() {} } },
  setTimeout() { return 1; },
  setInterval() { return 1; },
  clearTimeout() {},
  Intl,
  Map,
  Set,
  Date,
  Number,
  String,
  Boolean,
  Object,
  Array,
  Math,
  JSON
};
sandbox.window = sandbox;
sandbox.firebase.firestore.Query.prototype = { onSnapshot() {} };
sandbox.firebase.firestore.Query.prototype.__heraActiveCommesseOriginalOnSnapshot = sandbox.firebase.firestore.Query.prototype.onSnapshot;

vm.createContext(sandbox);
vm.runInContext(source, sandbox, { filename: "operational-import-repair.js" });

const testing = sandbox.INRETE_MODENA_AUGUST_2026?.testing;
assert.ok(testing, "API di test non disponibile");

const modena = commesseById.get("modena");
const ferrara = commesseById.get("ferrara");
const bologna = commesseById.get("bologna");

assert.equal(testing.isManagementForCommessa(modena), false, "Ferrara non deve essere identificata come Modena solo perché 28015001 contiene 28015");
assert.equal(testing.isManagementForCommessa(ferrara), true, "Ferrara deve riconoscere il proprio ID attivo");
assert.equal(testing.isManagementForCommessa(bologna), false, "Bologna non deve essere confusa con Ferrara o Modena");

const modenaSummary = {
  uniquePlants: 30,
  workRows: 42,
  doneWork: 4,
  pendingWork: 38,
  subtotalCompleted: 316.8
};

stats.innerHTML = "FERRARA ORIGINALE";
assert.equal(testing.applyManagementStats(modena, modenaSummary), false, "Il riepilogo Modena non deve essere applicato mentre è aperta Ferrara");
assert.equal(stats.innerHTML, "FERRARA ORIGINALE", "Le statistiche Ferrara non devono essere sovrascritte dai dati Modena");

sandbox.managementCommessaId = "modena";
meta.textContent = "INRETE MODENA • Cod. 28015";
assert.equal(testing.isManagementForCommessa(modena), true);
assert.equal(testing.applyManagementStats(modena, modenaSummary), true);
assert.match(stats.innerHTML, />30<\/b> impianti/);
assert.match(stats.innerHTML, />42<\/b> lavorazioni/);

sandbox.managementCommessaId = "";
meta.textContent = "INRETE FERRARA • Cod. 28015001";
assert.equal(testing.getManagementMetaCode(), "28015001");
assert.equal(testing.isManagementForCommessa(modena), false, "Il fallback sul codice deve usare uguaglianza esatta");
meta.textContent = "INRETE BOLOGNA • Cod. 28015002";
assert.equal(testing.isManagementForCommessa(modena), false, "28015002 non deve corrispondere a 28015");
meta.textContent = "INRETE MODENA • Cod. 28015";
assert.equal(testing.isManagementForCommessa(modena), true, "Il codice esatto 28015 deve corrispondere a Modena");

const ferraraOperational = Array.from({ length: 94 }, (_, index) => ({
  id: `op-${index + 1}`,
  idSap: `F${String(index + 1).padStart(3, "0")}`,
  denominazione: `FERRARA ${index + 1}`,
  comune: "FERRARA",
  stato: "FATTO",
  statoGenerale: "FATTO",
  done: true
}));
const ferraraPhysical = ferraraOperational.slice(0, 30).map((plant, index) => ({
  id: `physical-${index + 1}`,
  idSap: plant.idSap,
  denominazione: plant.denominazione,
  comune: plant.comune
}));
const ferraraWork = Array.from({ length: 42 }, (_, index) => ({
  id: `work-${index + 1}`,
  impiantoId: `physical-${(index % 30) + 1}`,
  stato: index < 4 ? "FATTO" : "DA FARE",
  totale: index < 4 ? 79.2 : null
}));

const ferraraSummary = testing.summarizeManagementData(ferraraOperational, ferraraPhysical, ferraraWork);
assert.equal(ferraraSummary.uniquePlants, 94, "La Gestione Ferrara deve usare i 94 impianti reali della raccolta operativa");
assert.equal(ferraraSummary.workRows, 42, "Le lavorazioni restano un conteggio separato dagli impianti");
assert.equal(ferraraSummary.doneWork, 4);
assert.equal(ferraraSummary.pendingWork, 38);
assert.equal(ferraraSummary.donePlants, 94, "Lo stato impianti operativo non deve essere degradato dai dati contabili parziali");

sandbox.managementCommessaId = "ferrara";
meta.textContent = "INRETE FERRARA • Cod. 28015001";
stats.innerHTML = "";
assert.equal(testing.renderManagementStatsForCommessa(ferrara, ferraraSummary), true);
assert.match(stats.innerHTML, />94<\/b> impianti/);
assert.match(stats.innerHTML, />42<\/b> lavorazioni/);
assert.match(stats.innerHTML, />4<\/b> lavorazioni fatte/);
assert.match(stats.innerHTML, />38<\/b> lavorazioni da fare/);

assert.doesNotMatch(source, /normalizeCode\(meta\.textContent\)\.includes\(CANONICAL_COMMESSA_CODE\)/);
assert.match(source, /if \(cached && options\.force !== true\) renderManagementStatsForCommessa/, "L'apertura della gestione non deve mostrare il riepilogo precedente prima della rilettura forzata");
assert.match(source, /const openedNow = isVisible && !managementScreenWasVisible/, "La riapertura della gestione deve invalidare il riepilogo conservato");
assert.match(source, /if \(canonical\?\.id === commessa\.id\)[\s\S]*applySummary\(canonical, summary\)/, "Il riepilogo aggiornato di INRETE Modena deve aggiornare anche totale e contatori della commessa");

assert.doesNotMatch(
  appSource,
  /if\s*\(\s*!ui\.commessaTargetSelect\.value\s*\)/,
  "Aprendo una commessa il selettore import non deve conservare una destinazione precedente"
);
assert.match(
  appSource,
  /if\s*\(ui\.commessaTargetSelect\)\s*\{\s*ui\.commessaTargetSelect\.value\s*=\s*id;/,
  "La destinazione import deve essere riallineata sempre alla commessa aperta"
);

const loadStart = accountingSource.indexOf("async function load(options={}){");
const loadEnd = accountingSource.indexOf("\n  const clean=", loadStart);
assert.ok(loadStart >= 0 && loadEnd > loadStart, "Caricamento contabile non trovato");
const loadSource = accountingSource.slice(loadStart, loadEnd);
assert.match(loadSource, /const requestId=Number\(options\.requestId\?\?state\.openRequestId\)/);
assert.match(loadSource, /const requestedRef=db\.collection\(getCommesseCollectionName\(\)\)\.doc\(requestedCommessaId\)/);
assert.match(loadSource, /if\(isStale\(\)\)return false;/);
assert.doesNotMatch(loadSource, /commRef\(\)\.collection\(/, "Il caricamento non deve rileggere la commessa globale mutabile");

const syncStart = accountingSource.indexOf("async function synchronizeOperationalModel(options={}){");
const syncEnd = accountingSource.indexOf("\n  async function verifyOperationalModel", syncStart);
assert.ok(syncStart >= 0 && syncEnd > syncStart, "Sincronizzazione operativa non trovata");
const syncSource = accountingSource.slice(syncStart, syncEnd);
assert.match(syncSource, /const workItems=state\.work\.map\(item=>\(\{\.\.\.item\}\)\)/);
assert.match(syncSource, /requestedRef\.collection\("impianti"\)/);
assert.match(syncSource, /commessaId:requestedCommessaId/);
assert.doesNotMatch(syncSource, /commRef\(\)/, "La sincronizzazione non deve cambiare riferimento durante gli await");
assert.doesNotMatch(syncSource, /commessaId:state\.commessa\.id/, "Nessun dato deve ricevere l’ID di una commessa cambiata");

assert.match(
  accountingSource,
  /const requestId=\+\+state\.openRequestId;state\.commessa=commessa;managementCommessaId=requestedCommessaId;/
);
assert.match(
  accountingSource,
  /return load\(\{commessa,commessaId:requestedCommessaId,requestId\}\);/
);

console.log("✅ Gestione commesse: isolamento esatto, destinazione import e concorrenza asincrona verificati.");

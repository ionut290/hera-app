#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync("operational-import-repair.js", "utf8");

const commesseById = new Map([
  ["modena", { id: "modena", nome: "INRETE MODENA", codice: "28015", impiantiCount: 30 }],
  ["ferrara", { id: "ferrara", nome: "INRETE FERRARA", codice: "28015001", impiantiCount: 94 }],
  ["bologna", { id: "bologna", nome: "INRETE BOLOGNA", codice: "28015002", impiantiCount: 120 }]
]);

const meta = { textContent: "INRETE FERRARA • Cod. 28015001" };
const stats = { innerHTML: "FERRARA ORIGINALE" };

const sandbox = {
  console: { info() {}, warn() {}, error() {} },
  addEventListener() {},
  document: {
    hidden: false,
    addEventListener() {},
    querySelector(selector) {
      if (selector === "#impianti-management-meta") return meta;
      if (selector === "#impianti-management-stats") return stats;
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

assert.doesNotMatch(source, /normalizeCode\(meta\.textContent\)\.includes\(CANONICAL_COMMESSA_CODE\)/);

console.log("✅ Isolamento commesse: Modena, Ferrara e Bologna mantengono statistiche indipendenti.");

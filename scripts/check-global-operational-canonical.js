#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync("commessa-stats-cache-optimizer.js", "utf8");

function originalCombine(items) {
  return Array.isArray(items) ? items.map((item) => ({ ...item })) : [];
}

const commesseById = new Map([
  ["ferrara", { id: "ferrara", nome: "INRETE FERRARA", workItemsModelVersion: 2, impiantiCount: 30, totalPlants: 30 }],
  ["legacy", { id: "legacy", nome: "LEGACY", workItemsModelVersion: 1, impiantiCount: 3 }]
]);

const sandbox = {
  console,
  window: {},
  localStorage: {
    data: new Map(),
    getItem(key) { return this.data.get(key) || null; },
    setItem(key, value) { this.data.set(key, value); },
    removeItem(key) { this.data.delete(key); },
    key(index) { return [...this.data.keys()][index] || null; },
    get length() { return this.data.size; }
  },
  currentUser: { uid: "u1" },
  selectedCommessaId: "ferrara",
  commesseById,
  impiantiByCommessaId: new Map(),
  commessaStatsById: new Map(),
  unsubscribeCommessaStats: new Map(),
  combineImpiantiForView: originalCombine,
  calculateImpiantiStats(items) {
    const combined = sandbox.combineImpiantiForView(items);
    return { total: combined.length, done: combined.filter((item) => item.done).length };
  },
  recalculateCommessaWorkSummaries() {},
  subscribeStatsForCommesse() {},
  getSubcommesse() { return []; },
  renderCommesseHomeList() {},
  renderCommesseManagementList() {},
  renderParentCommessaOverview() {},
  updateCommessaDashboard() {},
  getCommesseCollectionName() { return "commesse"; },
  db: {
    collection() {
      return {
        doc() {
          return {
            collection() {
              return {
                orderBy() { return this; },
                limit() { return this; },
                where() { return this; },
                async get() { return { empty: true, docs: [] }; },
                onSnapshot() { return () => {}; },
                doc() { return { get: async () => ({ exists: false }) }; }
              };
            }
          };
        }
      };
    }
  },
  setTimeout(fn) { fn(); return 1; },
  clearTimeout() {},
  Date,
  Map,
  Set,
  Number,
  String,
  Boolean,
  JSON,
  Object,
  Array,
  Math
};
sandbox.window = sandbox;
vm.createContext(sandbox);
vm.runInContext(source, sandbox, { filename: "commessa-stats-cache-optimizer.js" });

assert.ok(sandbox.HeraOperationalCanonicalView?.installed, "Filtro canonico operativo non installato");

const canonical = Array.from({ length: 30 }, (_, index) => ({
  id: `canonical-${index + 1}`,
  commessaId: "ferrara",
  physicalPlantId: `physical-${index + 1}`,
  migrationSourceId: `physical-${index + 1}`,
  idSap: String(3430000 + index),
  denominazione: `Impianto ${index + 1}`,
  numeroLavorazioni: index < 12 ? 2 : 1,
  numeroLavorazioniFatte: index < 4 ? (index < 12 ? 2 : 1) : 0,
  stato: index < 4 ? "FATTO" : "DA FARE",
  statoGenerale: index < 4 ? "FATTO" : "DA FARE",
  done: index < 4
}));

const legacyExtras = Array.from({ length: 64 }, (_, index) => ({
  id: `legacy-${index + 1}`,
  commessaId: "ferrara",
  idSap: String(9000000 + index),
  denominazione: `Vecchio impianto ${index + 1}`,
  stato: "FATTO",
  statoGenerale: "FATTO",
  done: true
}));

const rawFerrara = [...canonical, ...legacyExtras];
const filteredFerrara = sandbox.combineImpiantiForView(rawFerrara);
assert.equal(filteredFerrara.length, 30, "94 record operativi devono diventare i 30 impianti canonici correnti");
assert.equal(filteredFerrara.filter((item) => item.done).length, 4, "Gli stati FATTO devono provenire dai documenti canonici correnti");
assert.equal(filteredFerrara.some((item) => String(item.id).startsWith("legacy-")), false, "I record storici non canonici non devono apparire in mappa/lista");

const ferraraStats = sandbox.calculateImpiantiStats(rawFerrara);
assert.equal(ferraraStats.total, 30);
assert.equal(ferraraStats.done, 4);

const legacyOnly = [
  { id: "a", commessaId: "legacy", idSap: "1", done: false },
  { id: "b", commessaId: "legacy", idSap: "2", done: true },
  { id: "c", commessaId: "legacy", idSap: "3", done: false }
];
assert.equal(sandbox.combineImpiantiForView(legacyOnly).length, 3, "Le commesse legacy non devono perdere impianti");

const incompleteCanonical = canonical.slice(0, 10);
const incompleteRaw = [...incompleteCanonical, ...legacyOnly.map((item, index) => ({ ...item, commessaId: "ferrara", id: `fallback-${index}` }))];
commesseById.set("ferrara", { id: "ferrara", nome: "INRETE FERRARA", workItemsModelVersion: 1, impiantiCount: 30, totalPlants: 30 });
assert.equal(
  sandbox.combineImpiantiForView(incompleteRaw).length,
  incompleteRaw.length,
  "Una migrazione incompleta senza modello v2 non deve nascondere record legacy"
);

sandbox.localStorage.setItem("heraImpiantiPersistentCacheV1:u1:commesse:ferrara", "stale");
sandbox.HeraOperationalCanonicalView.clearLegacyCaches();
assert.equal(sandbox.localStorage.getItem("heraImpiantiPersistentCacheV1:u1:commesse:ferrara"), null, "La cache impianti precedente deve essere invalidata");

console.log("✅ Vista operativa canonica globale: 94→30, stati correnti preservati e commesse legacy protette.");

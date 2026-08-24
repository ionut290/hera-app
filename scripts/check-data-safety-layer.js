#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const { execFileSync } = require("node:child_process");
const vm = require("node:vm");

function read(path) {
  assert.ok(fs.existsSync(path), `File mancante: ${path}`);
  return fs.readFileSync(path, "utf8");
}

const safety = read("data-safety-layer.js");
const criticalBridge = read("critical-write-safety-bridge.js");
const durability = read("data-durability-runtime.js");
const updater = read("update-app-feature.js");
const sw = read("sw.js");
const pkg = JSON.parse(read("package.json"));

execFileSync(process.execPath, ["--check", "data-safety-layer.js"], { stdio: "inherit" });
execFileSync(process.execPath, ["--check", "critical-write-safety-bridge.js"], { stdio: "inherit" });

assert.match(safety, /window\.HeraDataSafety\s*=\s*\{/);
assert.match(safety, /createOperationId/);
assert.match(safety, /operationId/);
assert.match(safety, /registerSchema/);
assert.match(safety, /function migrate\(/);
assert.match(safety, /async function run\(/);
assert.match(safety, /before-write:/);
assert.match(safety, /after-write:/);
assert.match(safety, /installOfflineQueueObserver/);
assert.match(safety, /installOfflineSyncObserver/);
assert.match(safety, /enqueueOfflineMutation/);
assert.match(safety, /syncPendingOfflineMutations/);
assert.match(safety, /hera_data_safety_journal_v1/);
assert.match(safety, /hera_data_safety_completed_v1/);
assert.match(safety, /MAX_PAYLOAD_BYTES/);
assert.doesNotMatch(safety, /\.collection\s*\(/, "Il safety layer non deve scrivere direttamente su Firestore");
assert.doesNotMatch(safety, /localStorage\.clear\s*\(/, "Il safety layer non deve cancellare tutto il localStorage");
assert.doesNotMatch(safety, /indexedDB\.deleteDatabase\s*\(/, "Il safety layer non deve cancellare IndexedDB");

assert.match(criticalBridge, /window\.HeraCriticalWriteSafetyBridge\s*=\s*\{/);
assert.doesNotMatch(criticalBridge, /name:\s*"forceMoveImpiantoToFatti"/);
assert.doesNotMatch(criticalBridge, /name:\s*"markImpiantoDone"/);
assert.doesNotMatch(criticalBridge, /name:\s*"resetImpianto"/);
assert.match(criticalBridge, /deleteImpianto/);
assert.match(criticalBridge, /saveCommessaNote/);
assert.match(criticalBridge, /saveHoursReport/);
assert.match(criticalBridge, /saveSquadra/);
assert.match(criticalBridge, /safety\.run\(/);
assert.match(criticalBridge, /safety\.snapshot/);
assert.match(criticalBridge, /const dedupeIdentity = meta\.entityId \|\| operationId;/);
assert.match(criticalBridge, /const dedupeKey = `\$\{target\.name\}:\$\{target\.type\}:/);
assert.match(criticalBridge, /__heraCriticalWriteSafetyWrapped/);
assert.doesNotMatch(criticalBridge, /\.collection\s*\(/, "Il bridge non deve creare proprie scritture Firestore");
assert.doesNotMatch(criticalBridge, /localStorage\.clear\s*\(/, "Il bridge non deve cancellare dati locali");

assert.match(durability, /window\.HeraDataDurability/);
assert.match(durability, /async function snapshot\(/);
assert.match(updater, /data-safety-layer\.js\?v=20260819a/);
assert.match(updater, /ensureDataSafetyLayer/);
assert.match(updater, /critical-write-safety-bridge\.js\?v=20260824-oneclick1/);
assert.match(updater, /ensureCriticalWriteSafetyBridge/);
assert.match(sw, /\.\/data-safety-layer\.js\?v=20260819a/);
assert.match(sw, /"\/data-safety-layer\.js"/);
assert.doesNotMatch(sw, /client\.navigate\s*\(/, "L'attivazione del Service Worker non deve ricaricare forzatamente le finestre");

assert.equal(pkg.scripts["check:data-safety"], "node scripts/check-data-safety-layer.js");

async function checkNestedFattoFlowDoesNotDeadlock() {
  const storage = new Map();
  const sandbox = {
    console,
    Blob,
    crypto: globalThis.crypto,
    navigator: { onLine: true },
    localStorage: {
      getItem: (key) => storage.get(key) || null,
      setItem: (key, value) => storage.set(key, String(value))
    },
    CustomEvent: class {
      constructor(type, init = {}) {
        this.type = type;
        this.detail = init.detail;
      }
    },
    setTimeout: () => 0,
    clearTimeout: () => {},
    setInterval: () => 0,
    clearInterval: () => {},
    addEventListener: () => {},
    dispatchEvent: () => {},
    selectedCommessaId: "commessa-inrete"
  };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(safety, sandbox, { filename: "data-safety-layer.js" });
  vm.runInContext(`
    var markCalls = 0;
    async function markImpiantoDone() {
      markCalls += 1;
      return true;
    }
    async function forceMoveImpiantoToFatti(impianto, options = {}) {
      return Boolean(await markImpiantoDone(impianto, options));
    }
  `, sandbox);
  vm.runInContext(criticalBridge, sandbox, { filename: "critical-write-safety-bridge.js" });

  const timeout = new Promise((resolve) => setTimeout(() => resolve("timeout"), 250));
  const result = await Promise.race([
    sandbox.forceMoveImpiantoToFatti(
      { id: "impianto-6" },
      { commessaId: "commessa-inrete" }
    ),
    timeout
  ]);

  assert.equal(result, true, "FORZA/FATTO deve restare sul percorso diretto senza attese del bridge");
  assert.equal(sandbox.markCalls, 1, "La scrittura FATTO interna deve essere eseguita una sola volta");
  assert.equal(sandbox.HeraDataSafety.getState().inflightCount, 0, "La coda sicurezza deve liberarsi");
  assert.equal(sandbox.forceMoveImpiantoToFatti.__heraCriticalWriteSafetyWrapped, undefined, "Il bridge non deve avvolgere FATTO");
}

checkNestedFattoFlowDoesNotDeadlock()
  .then(() => console.log("✅ Data Safety Layer e bridge scritture critiche: controlli statici e deadlock FATTO superati."))
  .catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });

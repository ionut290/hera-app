#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const { execFileSync } = require("node:child_process");

function read(path) {
  assert.ok(fs.existsSync(path), `File mancante: ${path}`);
  return fs.readFileSync(path, "utf8");
}

const safety = read("data-safety-layer.js");
const durability = read("data-durability-runtime.js");
const updater = read("update-app-feature.js");
const sw = read("sw.js");
const pkg = JSON.parse(read("package.json"));

execFileSync(process.execPath, ["--check", "data-safety-layer.js"], { stdio: "inherit" });

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

assert.match(durability, /window\.HeraDataDurability/);
assert.match(durability, /async function snapshot\(/);
assert.match(updater, /data-safety-layer\.js\?v=20260819a/);
assert.match(updater, /ensureDataSafetyLayer/);
assert.match(sw, /\.\/data-safety-layer\.js\?v=20260819a/);
assert.match(sw, /"\/data-safety-layer\.js"/);
assert.doesNotMatch(sw, /client\.navigate\s*\(/, "L'attivazione del Service Worker non deve ricaricare forzatamente le finestre");

assert.equal(pkg.scripts["check:data-safety"], "node scripts/check-data-safety-layer.js");

console.log("✅ Data Safety Layer: controlli statici superati.");

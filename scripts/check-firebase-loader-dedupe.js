#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("firebase-config.js", "utf8");

assert.match(source, /function normalizeAssetPath\(/);
assert.match(source, /function findExistingScript\(src, dataName\)/);
assert.match(source, /Array\.from\(document\.scripts \|\| \[\]\)/);
assert.match(source, /normalizeAssetPath\(script\.src\) === wantedPath/);
assert.match(source, /const existing = findExistingScript\(src, dataName\)/);
assert.match(source, /if \(ready\?\.\(\)\)/);

for (const asset of [
  "firestore-operation-diagnostics.js",
  "firestore-diagnostics-dashboard-v4.js",
  "firestore-safe-optimizer.js",
  "firestore-inflight-read-coalescer.js",
  "shared-static-views.js",
  "active-commesse-first-boot-guard.js",
  "firestore-startup-cost-optimizer.js",
  "native-android-runtime.js",
  "admin-user-access-tools.js"
]) {
  assert.match(source, new RegExp(asset.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")), `Bootstrap mancante: ${asset}`);
}

console.log("✅ Bootstrap Firebase deduplica per stato, data attribute e src.");

#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("header-menu-runtime.js", "utf8");

assert.match(source, /function normalizeAssetPath\(/);
assert.match(source, /function findExistingScript\(/);
assert.match(source, /function loadScriptOnce\(/);
assert.match(source, /Array\.from\(document\.scripts \|\| \[\]\)/);
assert.match(source, /normalizeAssetPath\(script\.src\) === wantedPath/);

const manualAppendCount = (source.match(/document\.head\.appendChild\(script\)/g) || []).length;
assert.equal(manualAppendCount, 1, "I moduli JS del menu devono passare da un solo loader condiviso");

for (const asset of [
  "firestore-presence-cost-guard.js",
  "varga-branding.js",
  "operator-profile-feature.js",
  "global-archive-sync.js",
  "global-archive-new-commesse-fix.js",
  "preventivi-lazy-loader.js",
  "control-center-firestore-usage.js",
  "control-center-backup.js",
  "admin-console.js"
]) {
  assert.match(source, new RegExp(asset.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")), `Modulo mancante dal bootstrap: ${asset}`);
}

console.log("✅ Header/menu usa un solo loader JS con deduplicazione per src.");

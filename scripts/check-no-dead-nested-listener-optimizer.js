#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const sw = fs.readFileSync("sw.js", "utf8");
assert.doesNotMatch(sw, /firestore-nested-listener-optimizer\.js/);
assert.equal(fs.existsSync("firestore-nested-listener-optimizer.js"), false);
assert.equal(fs.existsSync("scripts/check-firestore-nested-listener-optimizer.js"), false);
const cacheMatch = sw.match(/CACHE_NAME\s*=\s*"varga-cantieri-shell-v(\d+)"/);
assert.ok(cacheMatch, "La shell PWA deve mantenere un nome cache versionato");
assert.ok(Number(cacheMatch[1]) >= 133, "La cache PWA non deve tornare a una versione precedente alla v133");
console.log("✅ Ottimizzatore listener morto rimosso anche dalla shell PWA.");

#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const sw = fs.readFileSync("sw.js", "utf8");
assert.doesNotMatch(sw, /firestore-nested-listener-optimizer\.js/);
assert.equal(fs.existsSync("firestore-nested-listener-optimizer.js"), false);
assert.equal(fs.existsSync("scripts/check-firestore-nested-listener-optimizer.js"), false);
assert.match(sw, /varga-cantieri-shell-v132/);
console.log("✅ Ottimizzatore listener morto rimosso anche dalla shell PWA.");

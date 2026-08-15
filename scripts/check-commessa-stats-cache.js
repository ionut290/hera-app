#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const code = fs.readFileSync("commessa-stats-cache-optimizer.js", "utf8");
const index = fs.readFileSync("index.html", "utf8");

assert.match(index, /commessa-stats-cache-optimizer\.js[^\n]*commessa-navigation-repair\.js/s, "Optimizer must load before navigation repair");
assert.match(code, /heraCommessaStatsCacheV1:/);
assert.match(code, /impiantoChangeIndex/);
assert.match(code, /where\("changedAt", ">", markerDate\)/);
assert.match(code, /orderBy\("changedAt", "asc"\)/);
assert.match(code, /\.doc\(impiantoId\)\s*;\s*const snap = await ref\.get\(\)/s);
assert.match(code, /fallbackFullLoad/);
assert.match(code, /collection\("impianti"\)\.get\(\)/);
assert.match(code, /readImpiantiCache/);
assert.match(code, /calculateImpiantiStats/);
assert.match(code, /recalculateCommessaWorkSummaries/);

assert.doesNotMatch(code, /\.set\(|\.update\(|\.delete\(|\.add\(/, "Optimizer must not write to Firestore");
assert.doesNotMatch(code, /FATTO|fattoVisualEvidence|WhatsApp|WHAZZUP|whazzup/i, "Protected FATTO/WhatsApp flows must stay outside this optimizer");

console.log("Commessa stats cache safety checks passed");

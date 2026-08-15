#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("functions/impianto-change-index.js", "utf8");
const main = fs.readFileSync("functions/main.js", "utf8");
const deploy = fs.readFileSync(".github/workflows/deploy-firebase-functions.yml", "utf8");

assert.match(source, /onDocumentWritten/);
assert.match(source, /document:\s*"commesse\/\{commessaId\}\/impianti\/\{impiantoId\}"/);
assert.match(source, /collection\("impiantoChangeIndex"\)/);
assert.match(source, /changedAt:\s*admin\.firestore\(\)\.FieldValue\.serverTimestamp\(\)/);
assert.match(source, /deleted:\s*!afterExists/);
assert.doesNotMatch(source, /collection\("impianti"\).*?\.set\(/s, "Il trigger non deve riscrivere gli impianti e non deve creare ricorsione");
assert.match(main, /require\("\.\/impianto-change-index"\)/);
assert.match(main, /impiantoChangeIndexFunctions/);
assert.match(deploy, /functions:syncImpiantoChangeIndex/);

console.log("✅ Change index impianti: trigger isolato, tombstone cancellazioni e deploy configurato.");
console.log("✅ Nessuna scrittura del trigger torna sulla collezione impianti.");

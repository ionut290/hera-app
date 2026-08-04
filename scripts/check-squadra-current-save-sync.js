#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const assert = require("node:assert/strict");

const source = fs.readFileSync("squadra-current-save-sync.js", "utf8");

assert.match(source, /currentComposition\(commessaId, dateKey\)/);
assert.match(source, /autofillSquadraForm = async function autofillCurrentSquadraFirst/);
assert.match(source, /Composizione salvata/);
assert.match(source, /squadreHistoryByDate\.set\(dateKey, history\)/);
assert.match(source, /showCurrentComposition\(commessaId, dateKey, composition\)/);
assert.doesNotMatch(source, /\.collection\(|\.get\(|\.onSnapshot\(|runFirestoreGetWithRetry/);

console.log("✅ Sincronizzazione vista squadra dopo salvataggio verificata.");

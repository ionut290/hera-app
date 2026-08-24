#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const immediate = fs.readFileSync("fatto-button-immediate.js", "utf8");
const config = fs.readFileSync("firebase-config.js", "utf8");
const operational = fs.readFileSync("operational-import-repair.js", "utf8");
const inreteCore = fs.readFileSync("inrete-work-items-v2.js", "utf8");

for (const source of [immediate, config, operational, inreteCore]) {
  assert.doesNotMatch(source, /PARZIALMENTE FATTO/, "Lo stato parziale non deve più esistere");
}
assert.doesNotMatch(immediate, /maybeHandlePartialInreteDone|loadInreteWorkContext|chooseWorkKinds|saveSelectedWorkItems/);
assert.doesNotMatch(immediate, /Cosa hai eseguito\?|Manutenzione ordinaria|Manutenzione straordinaria/);
assert.doesNotMatch(config, /fatto-ordinary-extraordinary-flow|data-fatto-ordinary-extraordinary/);
assert.match(immediate, /const operationTask = enqueue\(impianto/);
assert.doesNotMatch(immediate, /operation = await enqueue\(impianto/);
assert.match(immediate, /const result = await original\.call\(this, impianto, \.\.\.args\)/);
assert.doesNotMatch(immediate, /applyPermanentYellowFeedback|fattoImmediateDate|YELLOW_BORDER/);
assert.match(immediate, /Coda locale FATTO non disponibile; continuo con il flusso principale/);

console.log("✅ Gestione FATTO parziale INRETE completamente rimossa.");

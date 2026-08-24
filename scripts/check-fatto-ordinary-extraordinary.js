#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const immediate = fs.readFileSync("fatto-button-immediate.js", "utf8");
const config = fs.readFileSync("firebase-config.js", "utf8");

assert.equal(fs.existsSync("fatto-ordinary-extraordinary-flow.js"), false, "Il modulo separato deve essere eliminato");
assert.doesNotMatch(config, /HERA_FATTO_ORDINARY_EXTRAORDINARY_SRC|fatto-ordinary-extraordinary-flow/);
assert.doesNotMatch(immediate, /ORDINARIO|STRAORDINARIO|selectedKinds|inretePartialFatto/);
assert.doesNotMatch(immediate, /HeraFattoSync = Object\.freeze\(\{[^}]*maybeHandlePartialInreteDone/);
assert.match(immediate, /window\.HeraFattoSync = Object\.freeze\(\{ enqueue, processQueue, refreshStatus, list \}\)/);

console.log("✅ Selettore ordinario/straordinario eliminato; FATTO usa solo il flusso definitivo.");

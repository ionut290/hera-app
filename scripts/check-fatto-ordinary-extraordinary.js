#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("fatto-button-immediate.js", "utf8");
const config = fs.readFileSync("firebase-config.js", "utf8");

assert.doesNotMatch(
  config,
  /fatto-ordinary-extraordinary-flow\.js|HERA_FATTO_ORDINARY_EXTRAORDINARY_SRC|data-fatto-ordinary-extraordinary/,
  "La gestione FATTO ordinario/straordinario deve restare disattivata"
);
assert.doesNotMatch(
  source,
  /const partial = await maybeHandlePartialInreteDone\(impianto\)/,
  "Il wrapper principale non deve avviare la selezione delle lavorazioni"
);
assert.match(source, /window\.handleImpiantoWhatsAppClick = wrapped/);
assert.match(source, /const result = await original\.call\(this, impianto, \.\.\.args\)/);
assert.match(source, /enqueueSafely/);
assert.match(source, /processQueue/);

console.log("✅ Gestione ordinario/straordinario disattivata; FATTO standard protetto.");

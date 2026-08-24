#!/usr/bin/env node
const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("fatto-button-immediate.js", "utf8");
const config = fs.readFileSync("firebase-config.js", "utf8");

assert.doesNotMatch(
  source,
  /const partial = await maybeHandlePartialInreteDone\(impianto\)/,
  "Il pulsante FATTO non deve più essere intercettato dalla selezione ordinario/straordinario"
);
assert.doesNotMatch(
  source,
  /HeraFattoSync = Object\.freeze\(\{[^}]*maybeHandlePartialInreteDone/,
  "La selezione INRETE non deve essere esposta come parte del flusso FATTO attivo"
);
assert.doesNotMatch(
  config,
  /fatto-ordinary-extraordinary-flow\.js|HERA_FATTO_ORDINARY_EXTRAORDINARY_SRC|data-fatto-ordinary-extraordinary/,
  "Il modulo ordinario/straordinario non deve essere caricato"
);
assert.match(source, /FALLBACK_STORE_KEY = "heraFattoSyncOperationsFallbackV1"/);
assert.match(source, /async function enqueueSafely\(impianto, metadata = \{\}\)/);
assert.match(source, /const operation = await enqueueSafely\(impianto, \{/);
assert.match(source, /const result = await original\.call\(this, impianto, \.\.\.args\)/);
assert.ok(
  source.indexOf("const operation = await enqueueSafely(impianto, {") <
    source.indexOf("const result = await original.call(this, impianto, ...args)"),
  "La cassaforte deve precedere il flusso FATTO/Whazzup standard"
);

console.log("✅ INRETE usa il flusso FATTO/Whazzup standard con cassaforte.");

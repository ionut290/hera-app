#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("fatto-ordinary-extraordinary-flow.js", "utf8");
const config = fs.readFileSync("firebase-config.js", "utf8");

assert.match(source, /Cosa hai eseguito\?/);
assert.match(source, /input\.type = "checkbox"/);
assert.match(source, /Manutenzione ordinaria/);
assert.match(source, /Manutenzione straordinaria/);
assert.match(source, /Hai scelto solo la manutenzione ordinaria/);
assert.match(source, /Hai scelto solo la manutenzione straordinaria/);
assert.match(source, /TORNA INDIETRO/);
assert.match(source, /CONFERMA E INVIA/);
assert.match(source, /Intervento eseguito: manutenzione ordinaria/);
assert.match(source, /Intervento eseguito: manutenzione straordinaria/);
assert.match(source, /Interventi eseguiti: manutenzione ordinaria e straordinaria/);
assert.match(source, /selectedKinds\.includes\(getWorkItemKind\(entry\)\)/);
assert.match(source, /PARZIALMENTE FATTO/);
assert.match(source, /impianto\.done = allDone/);
assert.match(source, /sourceIds/);
assert.match(source, /processingPlants/);
assert.match(source, /if \(selectedKinds\.length === 2\) return \{ handled: false \}/);
assert.match(source, /!original\.__heraQueueWrapped/, "Il selettore deve installarsi sopra la coda FATTO stabile");
assert.ok(
  source.indexOf("const partial = await handleSelection(impianto)") < source.indexOf("return original.call(this, impianto, ...args)"),
  "La scelta ordinaria/straordinaria deve avvenire prima del FATTO definitivo"
);
assert.doesNotMatch(source, /IMPIANTO FATTO/);
assert.match(config, /fatto-ordinary-extraordinary-flow\.js\?v=20260814a/);

console.log("✅ FATTO ordinario/straordinario: controlli superati.");

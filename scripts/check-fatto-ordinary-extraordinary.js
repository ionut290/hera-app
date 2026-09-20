#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const immediate = fs.readFileSync("fatto-button-immediate.js", "utf8");
const config = fs.readFileSync("firebase-config.js", "utf8");
const a1Guard = fs.readFileSync("maintenance-price-code-display-guard.js", "utf8");

assert.equal(fs.existsSync("fatto-ordinary-extraordinary-flow.js"), false, "Il modulo separato deve essere eliminato");
assert.doesNotMatch(config, /HERA_FATTO_ORDINARY_EXTRAORDINARY_SRC|fatto-ordinary-extraordinary-flow/);
assert.doesNotMatch(immediate, /ORDINARIO|STRAORDINARIO|selectedKinds|inretePartialFatto/);
assert.doesNotMatch(immediate, /HeraFattoSync = Object\.freeze\(\{[^}]*maybeHandlePartialInreteDone/);
assert.match(immediate, /window\.HeraFattoSync = Object\.freeze\(\{ enqueue, processQueue, refreshStatus, list \}\)/);

const context = {
  document: {
    readyState: "complete",
    getElementById() {
      return null;
    },
    addEventListener() {}
  },
  MutationObserver: class MutationObserver {
    observe() {}
    disconnect() {}
  },
  requestAnimationFrame(callback) {
    callback();
  }
};
context.window = context;
vm.createContext(context);

vm.runInContext(`
  function splitLegacyCodes(value) {
    return Array.from(new Set(
      String(value || "")
        .toUpperCase()
        .split(/[^A-Z0-9]+/)
        .map((code) => code.trim())
        .filter(Boolean)
    ));
  }

  function hasOrdinario(value) {
    const codes = splitLegacyCodes(value);
    return codes.includes("A11") || codes.includes("A12");
  }

  function hasStraordinario(value) {
    const codes = splitLegacyCodes(value);
    return codes.length > 0 && codes.some((code) => code !== "A11" && code !== "A12");
  }

  function buildProtectedStyleWhatsAppLabel(impianto) {
    const isOnlyOrdinaria = hasOrdinario(impianto.codicePrezzo)
      && !hasStraordinario(impianto.codicePrezzo);
    return isOnlyOrdinaria
      ? "Manutenzione ordinaria eseguita"
      : "Manutenzione ordinaria + straordinaria eseguita";
  }
`, context);

assert.equal(
  vm.runInContext('buildProtectedStyleWhatsAppLabel({ codicePrezzo: "A1" })', context),
  "Manutenzione ordinaria + straordinaria eseguita",
  "Il test deve riprodurre il comportamento precedente errato"
);

vm.runInContext(a1Guard, context, { filename: "maintenance-price-code-display-guard.js" });

assert.equal(context.hasOrdinario("A1"), true, "A1 deve essere ordinario");
assert.equal(context.hasStraordinario("A1"), false, "A1 non deve essere straordinario");
assert.equal(context.hasOrdinario("A11"), true, "A11 deve restare ordinario");
assert.equal(context.hasStraordinario("A12"), false, "A12 deve restare non straordinario");
assert.equal(context.hasOrdinario("A1 | B4"), true, "A1 deve restare ordinario anche con un altro codice");
assert.equal(context.hasStraordinario("A1 | B4"), true, "B4 deve mantenere la componente straordinaria");
assert.equal(context.hasOrdinario("B4"), false, "B4 non deve diventare ordinario");
assert.equal(context.hasStraordinario("B4"), true, "B4 deve restare straordinario");
assert.equal(
  vm.runInContext('buildProtectedStyleWhatsAppLabel({ codicePrezzo: "A1" })', context),
  "Manutenzione ordinaria eseguita",
  "Il messaggio Whazzup deve classificare A1 come ordinario"
);
assert.equal(
  vm.runInContext('buildProtectedStyleWhatsAppLabel({ codicePrezzo: "A1 | B4" })', context),
  "Manutenzione ordinaria + straordinaria eseguita",
  "Un vero codice aggiuntivo straordinario deve restare visibile"
);

console.log("✅ Selettore ordinario/straordinario eliminato; FATTO usa solo il flusso definitivo.");
console.log("✅ A1/A11/A12 classificati come ordinari anche nel messaggio Whazzup; gli altri codici restano invariati.");

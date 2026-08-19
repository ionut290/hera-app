#!/usr/bin/env node
"use strict";

const fs = require("node:fs");

function replaceOnce(source, before, after, label) {
  if (source.includes(after)) return source;
  const first = source.indexOf(before);
  if (first < 0) throw new Error(`Blocco non trovato: ${label}`);
  if (source.indexOf(before, first + before.length) >= 0) {
    throw new Error(`Blocco non univoco: ${label}`);
  }
  return source.slice(0, first) + after + source.slice(first + before.length);
}

const runtimePath = "operational-import-repair.js";
let runtime = fs.readFileSync(runtimePath, "utf8");

runtime = replaceOnce(
  runtime,
  `      attivitaLabel: "ORDINARIO",\n      sourceDataset: SOURCE_DATASET,`,
  `      attivitaLabel: "ORDINARIO",\n      localFallback: true,\n      sourceDataset: SOURCE_DATASET,`,
  "marca gli impianti generati come fallback locale"
);

runtime = replaceOnce(
  runtime,
  `        mergeObjectByPriority(group.item, item, priority);\n        if (priority === 0 && !group.fallback) group.fallback = { ...item };`,
  `        const effectivePriority = item.localFallback === true ? 0 : priority;\n        mergeObjectByPriority(group.item, item, effectivePriority);\n        if (priority === 0 && !group.fallback) group.fallback = { ...item };`,
  "calcolo priorità fallback"
);

runtime = replaceOnce(
  runtime,
  `          if (priority > group.kindPriority) {\n            group.kinds = new Set(detected);\n            group.kindPriority = priority;\n          } else if (priority === group.kindPriority) {`,
  `          if (effectivePriority > group.kindPriority) {\n            group.kinds = new Set(detected);\n            group.kindPriority = effectivePriority;\n          } else if (effectivePriority === group.kindPriority) {`,
  "priorità tipologia lavorazione"
);

runtime = replaceOnce(
  runtime,
  `    setIfDifferent("sourceDataset", SOURCE_DATASET);`,
  `    setIfDifferent("localFallback", false);\n    setIfDifferent("sourceDataset", SOURCE_DATASET);`,
  "rimozione flag fallback dai documenti Firestore"
);

runtime = replaceOnce(
  runtime,
  `  function buildNewPlantData(fallback, aggregate = null) {\n    const data = { ...fallback };`,
  `  function buildNewPlantData(fallback, aggregate = null) {\n    const data = { ...fallback, localFallback: false };`,
  "nuovi documenti Firestore non fallback"
);

fs.writeFileSync(runtimePath, runtime);

const testPath = "scripts/check-commesse-impianti-integrity.js";
let test = fs.readFileSync(testPath, "utf8");

const raceTest = `\nconst fallbackSnapshot = api.plants.map((plant) => ({ ...plant }));\nmerged = api.mergePlants(api.plants, [{\n  ...canonical,\n  localFallback: false,\n  latitudine: 44.60001,\n  longitudine: 11.05001,\n  gpsY: 44.60001,\n  gpsX: 11.05001,\n  stato: "FATTO",\n  statoGenerale: "FATTO",\n  done: true,\n  updatedAt: { seconds: 1, nanoseconds: 0 }\n}], fallbackSnapshot);\ntarget = merged.find((plant) => plant.id === canonical.id);\nassert.equal(target.statoGenerale, "FATTO", "Il fallback corrente non deve prevalere sul dato Firestore");\nassert.equal(target.done, true);\nassert.equal(target.latitudine, 44.60001);\nassert.equal(target.longitudine, 11.05001);\nassert.equal(target.localFallback, false);\n`;

test = replaceOnce(
  test,
  `assert.equal(target.longitudine, 11.05001);\n\napi.testing.setWorkItems([{`,
  `assert.equal(target.longitudine, 11.05001);\n${raceTest}\napi.testing.setWorkItems([{`,
  "regressione gara fallback Firestore"
);

test = replaceOnce(
  test,
  `assert.equal(Object.prototype.hasOwnProperty.call(safePatch, "done"), false);\n\nconst recoveryPatch`,
  `assert.equal(Object.prototype.hasOwnProperty.call(safePatch, "done"), false);\nassert.equal(safePatch.localFallback, false, "Il documento Firestore non deve restare marcato come fallback locale");\n\nconst recoveryPatch`,
  "verifica flag Firestore"
);

fs.writeFileSync(testPath, test);
console.log("Correzione gara fallback commesse-impianti applicata.");

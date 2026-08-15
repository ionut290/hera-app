#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const inputPath = path.join(ROOT, "docs", "phase30-module-audit.json");
const report = JSON.parse(fs.readFileSync(inputPath, "utf8"));

function summarizeNamed(items, pattern) {
  const selected = items.filter((item) => pattern.test(item.name));
  const names = new Set(selected.map((item) => item.name));
  return {
    functions: selected.length,
    bytes: selected.reduce((sum, item) => sum + item.bytes, 0),
    names: [...names].sort(),
    crossCalls: [...new Set(selected.flatMap((item) => item.topLevelCalls.filter((name) => !names.has(name))))].sort(),
    globalDependencies: [...new Set(selected.flatMap((item) => item.globalDeps))].sort(),
    domFunctions: selected.filter((item) => item.touchesDom).map((item) => item.name).sort(),
    firebaseFunctions: selected.filter((item) => item.touchesFirebase).map((item) => item.name).sort(),
    asyncFunctions: selected.filter((item) => item.async).map((item) => item.name).sort()
  };
}

const result = {
  atexNamed: summarizeNamed(report.clusters.atex || [], /atex/i),
  worklimateNamed: summarizeNamed(report.clusters.worklimate || [], /worklimate/i)
};

fs.writeFileSync(
  path.join(ROOT, "docs", "phase30-module-selection.json"),
  JSON.stringify(result, null, 2) + "\n",
  "utf8"
);
console.log(JSON.stringify(result, null, 2));

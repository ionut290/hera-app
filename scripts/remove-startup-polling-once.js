#!/usr/bin/env node
"use strict";

const fs = require("node:fs");

function replaceOnce(path, before, after) {
  const source = fs.readFileSync(path, "utf8");
  if (!source.includes(before)) throw new Error(`Blocco atteso non trovato in ${path}`);
  if (source.indexOf(before) !== source.lastIndexOf(before)) throw new Error(`Blocco ambiguo in ${path}`);
  fs.writeFileSync(path, source.replace(before, after));
}

replaceOnce(
  "active-commesse-first-boot-guard.js",
  `  if (!install()) {\n    let attempts = 0;\n    const timer = window.setInterval(() => {\n      attempts += 1;\n      if (install() || attempts >= 100) window.clearInterval(timer);\n      window[GLOBAL].installed = state.installed;\n    }, 25);\n  }\n  window[GLOBAL].installed = state.installed;`,
  `  install();\n  window[GLOBAL].installed = state.installed;`
);

replaceOnce(
  "firestore-startup-cost-optimizer.js",
  `  installLazyTriggers();\n  let attempts = 0;\n  const timer = setInterval(() => {\n    attempts += 1;\n    if (install() || attempts >= 200) clearInterval(timer);\n    api.installed = state.installed;\n  }, 25);\n  install();\n  api.installed = state.installed;`,
  `  installLazyTriggers();\n  install();\n  api.installed = state.installed;`
);

console.log("Polling tecnico di bootstrap rimosso dalle due guardie.");

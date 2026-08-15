#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const acorn = require("acorn");

const targetNames = [
  "getFerieEligibleOperators",
  "refreshFerieProgrammazioneUi",
  "refreshFerieOperatorOptions",
  "saveFerieCollega",
  "computeDayStats",
  "renderFerieList",
  "renderFerieDisponibilitaCalendar"
];

const appPath = "app.js";
const modulePath = "app-availability.js";
const indexPath = "index.html";
const swPath = "sw.js";

let app = fs.readFileSync(appPath, "utf8");
const ast = acorn.parse(app, { ecmaVersion: "latest", sourceType: "script", ranges: true });
const found = new Map();
for (const node of ast.body) {
  if (node.type === "FunctionDeclaration" && node.id?.name && targetNames.includes(node.id.name)) {
    found.set(node.id.name, node);
  }
}
const missing = targetNames.filter((name) => !found.has(name));
if (missing.length) throw new Error(`Funzioni disponibilità mancanti: ${missing.join(", ")}`);

const ordered = targetNames.map((name) => found.get(name)).sort((a, b) => a.start - b.start);
const functionsSource = ordered.map((node) => app.slice(node.start, node.end)).join("\n\n");
for (const node of [...ordered].sort((a, b) => b.start - a.start)) {
  app = app.slice(0, node.start) + app.slice(node.end);
}
app = app.replace(/\n{4,}/g, "\n\n\n");
fs.writeFileSync(appPath, app);

const moduleSource = `/* Modulo Ferie / Disponibilità personale - estratto dal core senza modifiche funzionali. */\n(function (global) {\n  "use strict";\n\n${functionsSource}\n\n  Object.assign(global, {\n${targetNames.map((name) => `    ${name}`).join(",\n")}\n  });\n  global.VargaAvailabilityModule = Object.freeze({\n    functions: Object.freeze(${JSON.stringify(targetNames)})\n  });\n})(globalThis);\n`;
fs.writeFileSync(modulePath, moduleSource);

let index = fs.readFileSync(indexPath, "utf8");
if (!index.includes('src="app-availability.js')) {
  const appScriptRe = /(<script[^>]+src=["'][^"']*app\.js(?:\?[^"']*)?["'][^>]*><\/script>)/i;
  if (!appScriptRe.test(index)) throw new Error("Script app.js non trovato in index.html");
  index = index.replace(appScriptRe, `<script src="app-availability.js"></script>\n  $1`);
  fs.writeFileSync(indexPath, index);
}

let sw = fs.readFileSync(swPath, "utf8");
if (!sw.includes("./app-availability.js")) {
  const appAssetRe = /(\s*["']\.\/app\.js["'],?)/;
  if (!appAssetRe.test(sw)) throw new Error("Asset ./app.js non trovato in sw.js");
  sw = sw.replace(appAssetRe, `\n  "./app-availability.js",$1`);
  fs.writeFileSync(swPath, sw);
}

console.log(`Estratte ${targetNames.length} funzioni in ${modulePath}.`);

#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");

const ROOT = path.resolve(__dirname, "..");
const APP = path.join(ROOT, "app.js");
const INDEX = path.join(ROOT, "index.html");
const SW = path.join(ROOT, "sw.js");
const OUT = path.join(ROOT, "app-worklimate.js");

const TARGETS = [
  "bindHomeWorklimateButton",
  "buildHomeWorklimateButton",
  "formatWorklimateTemperature",
  "getHomeWorklimateRiskLevel",
  "getMostSevereWorklimateRisk",
  "getSquadraWorklimateCodeLineMarkup",
  "getWorklimateAverageTemperature",
  "getWorklimateContextForCommessa",
  "getWorklimateContextForCommessaAtNoon",
  "getWorklimateRiskForDateHour",
  "loadWorklimateRiskCacheBackground",
  "normalizeWorklimateLevel",
  "normalizeWorklimateRiskDoc",
  "openHomeWorklimateBoard",
  "openSquadraWorklimateSafety"
];

let source = fs.readFileSync(APP, "utf8");
const ast = acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
const found = ast.body.filter((node) => node.type === "FunctionDeclaration" && TARGETS.includes(node.id?.name));
const foundNames = found.map((node) => node.id.name).sort();
const missing = TARGETS.filter((name) => !foundNames.includes(name));
if (missing.length) throw new Error(`Funzioni Worklimate mancanti: ${missing.join(", ")}`);
if (found.length !== TARGETS.length) throw new Error(`Attese ${TARGETS.length} funzioni, trovate ${found.length}`);

const parts = [
  '"use strict";',
  '(function installVargaWorklimateModule(global) {',
  '  if (global.VargaWorklimateModule) return;',
  '  const api = {};'
];

for (const fn of found.sort((a, b) => a.start - b.start)) {
  const text = source.slice(fn.start, fn.end);
  parts.push(text.split("\n").map((line) => "  " + line).join("\n"));
  parts.push(`  api.${fn.id.name} = ${fn.id.name};`);
}
parts.push('  Object.assign(global, api);');
parts.push('  global.VargaWorklimateModule = Object.freeze({ ...api });');
parts.push('})(window);');
parts.push('');

const moduleSource = parts.join("\n");
acorn.parse(moduleSource, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
fs.writeFileSync(OUT, moduleSource, "utf8");

const ranges = found.map((fn) => ({ start: fn.start, end: fn.end })).sort((a, b) => b.start - a.start);
for (const range of ranges) {
  let end = range.end;
  while (end < source.length && /[ \t]/.test(source[end])) end += 1;
  if (source[end] === "\r") end += 1;
  if (source[end] === "\n") end += 1;
  source = source.slice(0, range.start) + source.slice(end);
}
acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
fs.writeFileSync(APP, source, "utf8");

let index = fs.readFileSync(INDEX, "utf8");
if (!index.includes("app-worklimate.js")) {
  const appTag = index.match(/<script[^>]+src=["'][^"']*app\.js[^"']*["'][^>]*><\/script>/i);
  if (!appTag) throw new Error("Tag app.js non trovato in index.html");
  index = index.replace(appTag[0], `<script src="app-worklimate.js?v=20260815-mod1"></script>\n${appTag[0]}`);
  fs.writeFileSync(INDEX, index, "utf8");
}

let sw = fs.readFileSync(SW, "utf8");
if (!sw.includes("app-worklimate.js")) {
  const marker = /(["']\.\/app\.js[^"']*["'],?)/;
  if (!marker.test(sw)) throw new Error("app.js non trovato nella shell del Service Worker");
  sw = sw.replace(marker, `"./app-worklimate.js?v=20260815-mod1",\n  $1`);
  fs.writeFileSync(SW, sw, "utf8");
}

console.log(`Estratto modulo Worklimate: ${TARGETS.length} funzioni, ${moduleSource.length} byte.`);

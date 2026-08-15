#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");

const ROOT = path.resolve(__dirname, "..");
const APP = path.join(ROOT, "app.js");
const INDEX = path.join(ROOT, "index.html");
const SW = path.join(ROOT, "sw.js");
const OUT = path.join(ROOT, "app-atex.js");

const TARGETS = [
  "buildAtexChecklist",
  "buildAtexContactCard",
  "buildAtexEmergencyContactsSection",
  "buildAtexImageCard",
  "buildAtexList",
  "closeAtexProcedurePage",
  "formatAtexDateValue",
  "formatAtexTimeValue",
  "getAtexClientContactsForCommessa",
  "getAtexIllustrationSvg",
  "getAtexProcedureImpiantoByKey",
  "getAtexWhatsappText",
  "getCurrentAtexProcedureContext",
  "getSelectedAtexCommessaName",
  "handleAtexProcedureButtonClick",
  "handleAtexProcedureContentClick",
  "hasImpiantoAtexFlag",
  "isTruthyAtexFlag",
  "normalizeAtexCommessaMatchValue",
  "openAtexProcedurePage",
  "renderAtexProcedurePage",
  "saveAtexProcedureForm",
  "shouldShowAtexButtonForImpianto",
  "valueContainsAtex"
];

let source = fs.readFileSync(APP, "utf8");
const ast = acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
const found = ast.body.filter((node) => node.type === "FunctionDeclaration" && TARGETS.includes(node.id?.name));
const names = found.map((node) => node.id.name).sort();
const missing = TARGETS.filter((name) => !names.includes(name));
if (missing.length) throw new Error(`Funzioni ATEX mancanti: ${missing.join(", ")}`);
if (found.length !== TARGETS.length) throw new Error(`Attese ${TARGETS.length} funzioni, trovate ${found.length}`);

const parts = [
  '"use strict";',
  '(function installVargaAtexModule(global) {',
  '  if (global.VargaAtexModule) return;',
  '  const api = {};'
];
for (const fn of found.sort((a, b) => a.start - b.start)) {
  const text = source.slice(fn.start, fn.end);
  parts.push(text.split("\n").map((line) => "  " + line).join("\n"));
  parts.push(`  api.${fn.id.name} = ${fn.id.name};`);
}
parts.push('  Object.assign(global, api);');
parts.push('  global.VargaAtexModule = Object.freeze({ ...api });');
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
if (!index.includes("app-atex.js")) {
  const appTag = index.match(/<script[^>]+src=["'][^"']*app\.js[^"']*["'][^>]*><\/script>/i);
  if (!appTag) throw new Error("Tag app.js non trovato in index.html");
  index = index.replace(appTag[0], `<script src="app-atex.js?v=20260815-mod1"></script>\n${appTag[0]}`);
  fs.writeFileSync(INDEX, index, "utf8");
}

let sw = fs.readFileSync(SW, "utf8");
if (!sw.includes("app-atex.js")) {
  const marker = /(["']\.\/app\.js[^"']*["'],?)/;
  if (!marker.test(sw)) throw new Error("app.js non trovato nella shell PWA");
  sw = sw.replace(marker, `"./app-atex.js?v=20260815-mod1",\n  $1`);
  fs.writeFileSync(SW, sw, "utf8");
}

console.log(`Estratto modulo ATEX: ${TARGETS.length} funzioni, ${moduleSource.length} byte.`);

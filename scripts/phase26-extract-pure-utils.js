#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");

const ROOT = path.resolve(__dirname, "..");
const APP = path.join(ROOT, "app.js");
const INDEX = path.join(ROOT, "index.html");
const SW = path.join(ROOT, "sw.js");
const OUT = path.join(ROOT, "app-pure-utils.js");

const BUILTINS = new Set([
  "String","Number","Boolean","Array","Object","Math","Date","JSON","RegExp","Set","Map",
  "parseInt","parseFloat","isNaN","isFinite","encodeURIComponent","decodeURIComponent",
  "encodeURI","decodeURI","Intl","BigInt","undefined","NaN","Infinity"
]);
const FORBIDDEN_NAMES = /(?:fatto|whazz|whats|firebase|firestore|auth|login|squad|commess|impiant|ore|gps|map|weather|meteo|offline|sync|admin|user|photo|foto|document|window|ui)/i;

function walk(node, visitor, parent = null) {
  if (!node || typeof node !== "object") return;
  visitor(node, parent);
  for (const [key, value] of Object.entries(node)) {
    if (["start", "end", "loc"].includes(key)) continue;
    if (Array.isArray(value)) value.forEach((child) => child && typeof child === "object" && walk(child, visitor, node));
    else if (value && typeof value === "object" && value.type) walk(value, visitor, node);
  }
}

function declaredNames(fn) {
  const names = new Set([fn.id.name]);
  for (const p of fn.params) walk(p, (n) => { if (n.type === "Identifier") names.add(n.name); });
  walk(fn.body, (n) => {
    if (n.type === "VariableDeclarator") walk(n.id, (x) => { if (x.type === "Identifier") names.add(x.name); });
    if ((n.type === "FunctionDeclaration" || n.type === "ClassDeclaration") && n.id) names.add(n.id.name);
    if (n.type === "CatchClause" && n.param) walk(n.param, (x) => { if (x.type === "Identifier") names.add(x.name); });
  });
  return names;
}

function isReferenceIdentifier(node, parent) {
  if (!parent) return false;
  if (parent.type === "MemberExpression" && parent.property === node && !parent.computed) return false;
  if (parent.type === "Property" && parent.key === node && !parent.computed && !parent.shorthand) return false;
  if ((parent.type === "FunctionDeclaration" || parent.type === "FunctionExpression") && parent.id === node) return false;
  if (parent.type === "VariableDeclarator" && parent.id === node) return false;
  if (parent.type === "LabeledStatement" || parent.type === "BreakStatement" || parent.type === "ContinueStatement") return false;
  return true;
}

function isPureCandidate(fn, source) {
  const name = fn.id?.name || "";
  if (!name || FORBIDDEN_NAMES.test(name)) return false;
  const text = source.slice(fn.start, fn.end);
  if (text.length < 40 || text.length > 1800) return false;
  if (/\b(?:window|document|navigator|localStorage|sessionStorage|firebase|fetch|console|setTimeout|setInterval)\b/.test(text)) return false;
  const declared = declaredNames(fn);
  let safe = true;
  walk(fn.body, (node, parent) => {
    if (!safe) return;
    if (node.type === "ThisExpression" || node.type === "AwaitExpression" || node.type === "YieldExpression") { safe = false; return; }
    if (node.type === "Identifier" && isReferenceIdentifier(node, parent)) {
      if (!declared.has(node.name) && !BUILTINS.has(node.name)) safe = false;
    }
  });
  return safe;
}

let source = fs.readFileSync(APP, "utf8");
const ast = acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
const topFns = ast.body.filter((n) => n.type === "FunctionDeclaration");
const candidates = topFns.filter((fn) => isPureCandidate(fn, source));
if (!candidates.length) throw new Error("Nessuna funzione pura sicura trovata: nessuna modifica applicata.");

const selected = candidates.sort((a,b) => (a.end-a.start)-(b.end-b.start)).slice(0,4);
const names = selected.map((fn) => fn.id.name);

const moduleParts = [
  '"use strict";',
  '(function exposeVargaPureUtils(global) {',
  '  const api = {};'
];
for (const fn of selected) {
  const text = source.slice(fn.start, fn.end);
  moduleParts.push(text.split("\n").map((line) => "  " + line).join("\n"));
  moduleParts.push(`  api.${fn.id.name} = ${fn.id.name};`);
}
moduleParts.push('  Object.assign(global, api);');
moduleParts.push('  global.VargaPureUtils = Object.freeze({ ...(global.VargaPureUtils || {}), ...api });');
moduleParts.push('})(window);');
moduleParts.push('');
fs.writeFileSync(OUT, moduleParts.join("\n"), "utf8");

const ranges = selected.map((fn) => ({start:fn.start,end:fn.end})).sort((a,b)=>b.start-a.start);
for (const r of ranges) {
  let end = r.end;
  while (end < source.length && /[ \t]/.test(source[end])) end++;
  if (source[end] === "\r") end++;
  if (source[end] === "\n") end++;
  source = source.slice(0,r.start) + source.slice(end);
}
acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
fs.writeFileSync(APP, source, "utf8");

let index = fs.readFileSync(INDEX, "utf8");
if (!index.includes("app-pure-utils.js")) {
  const appTag = /<script[^>]+src=["'][^"']*app\.js[^"']*["'][^>]*><\/script>/i;
  const match = index.match(appTag);
  if (!match) throw new Error("Tag app.js non trovato in index.html");
  index = index.replace(match[0], `<script src="app-pure-utils.js?v=20260815-mod1"></script>\n${match[0]}`);
  fs.writeFileSync(INDEX, index, "utf8");
}

let sw = fs.readFileSync(SW, "utf8");
if (!sw.includes("app-pure-utils.js")) {
  const marker = /(["']\.\/app\.js[^"']*["'],?)/;
  if (marker.test(sw)) sw = sw.replace(marker, `"./app-pure-utils.js?v=20260815-mod1",\n  $1`);
  else console.warn("AVVISO: app.js non trovato nella precache SW; modulo non aggiunto automaticamente alla shell.");
  fs.writeFileSync(SW, sw, "utf8");
}

console.log(`Estratte ${names.length} funzioni pure in app-pure-utils.js: ${names.join(", ")}`);

#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");

const ROOT = path.resolve(__dirname, "..");
const APP = path.join(ROOT, "app.js");
const OUT = path.join(ROOT, "app-pure-utils.js");

const BUILTINS = new Set([
  "String","Number","Boolean","Array","Object","Math","Date","JSON","RegExp","Set","Map",
  "parseInt","parseFloat","isNaN","isFinite","encodeURIComponent","decodeURIComponent",
  "encodeURI","decodeURI","Intl","BigInt","undefined","NaN","Infinity"
]);
const FORBIDDEN_NAMES = /(?:fatto|whazz|whats|firebase|firestore|auth|login|squad|commess|impiant|ore|gps|map|weather|meteo|offline|sync|admin|user|photo|foto|document|window|ui|operator|position|corso|worklimate|phone|mezzo|minute)/i;

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
  if (text.length < 40 || text.length > 1400) return false;
  if (/\b(?:window|document|navigator|localStorage|sessionStorage|firebase|fetch|console|setTimeout|setInterval|location|history)\b/.test(text)) return false;
  const declared = declaredNames(fn);
  let safe = true;
  walk(fn.body, (node, parent) => {
    if (!safe) return;
    if (["ThisExpression","AwaitExpression","YieldExpression","NewExpression"].includes(node.type)) { safe = false; return; }
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
if (!candidates.length) throw new Error("Nessuna ulteriore utility pura sicura trovata.");
const selected = candidates.sort((a,b) => (a.end-a.start)-(b.end-b.start)).slice(0,4);
const names = selected.map((fn) => fn.id.name);

let utilSource = fs.readFileSync(OUT, "utf8");
const extra = ["", "// Phase 28: pure utilities extracted from app.js", "(function exposePhase28PureUtils(global) {", "  const api = {};"];
for (const fn of selected) {
  const text = source.slice(fn.start, fn.end);
  extra.push(text.split("\n").map((line) => "  " + line).join("\n"));
  extra.push(`  api.${fn.id.name} = ${fn.id.name};`);
}
extra.push("  Object.assign(global, api);");
extra.push("  global.VargaPureUtils = Object.freeze({ ...(global.VargaPureUtils || {}), ...api });");
extra.push("})(window);");
extra.push("");
utilSource = utilSource.trimEnd() + "\n" + extra.join("\n");
acorn.parse(utilSource, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
fs.writeFileSync(OUT, utilSource, "utf8");

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
console.log(`Fase 28: estratte ${names.length} utility pure: ${names.join(", ")}`);

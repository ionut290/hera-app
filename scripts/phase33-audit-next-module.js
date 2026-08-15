#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");

const ROOT = path.resolve(__dirname, "..");
const source = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
const ast = acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
const topFns = ast.body.filter((node) => node.type === "FunctionDeclaration" && node.id?.name);
const topNames = new Set(topFns.map((fn) => fn.id.name));

function walk(node, visitor, parent = null) {
  if (!node || typeof node !== "object") return;
  visitor(node, parent);
  for (const [key, value] of Object.entries(node)) {
    if (["start", "end", "loc"].includes(key)) continue;
    if (Array.isArray(value)) value.forEach((child) => child?.type && walk(child, visitor, node));
    else if (value?.type) walk(value, visitor, node);
  }
}

function declaredNames(fn) {
  const names = new Set([fn.id?.name].filter(Boolean));
  for (const param of fn.params || []) walk(param, (node) => { if (node.type === "Identifier") names.add(node.name); });
  walk(fn.body, (node) => {
    if (node.type === "VariableDeclarator") walk(node.id, (child) => { if (child.type === "Identifier") names.add(child.name); });
    if ((node.type === "FunctionDeclaration" || node.type === "ClassDeclaration") && node.id) names.add(node.id.name);
    if (node.type === "CatchClause" && node.param) walk(node.param, (child) => { if (child.type === "Identifier") names.add(child.name); });
  });
  return names;
}

function isReference(node, parent) {
  if (!parent) return false;
  if (parent.type === "MemberExpression" && parent.property === node && !parent.computed) return false;
  if (parent.type === "Property" && parent.key === node && !parent.computed && !parent.shorthand) return false;
  if ((parent.type === "FunctionDeclaration" || parent.type === "FunctionExpression") && parent.id === node) return false;
  if (parent.type === "VariableDeclarator" && parent.id === node) return false;
  return true;
}

function analyze(fn) {
  const declared = declaredNames(fn);
  const calls = new Set();
  const globals = new Set();
  const text = source.slice(fn.start, fn.end);
  walk(fn.body, (node, parent) => {
    if (node.type === "CallExpression" && node.callee?.type === "Identifier") calls.add(node.callee.name);
    if (node.type === "Identifier" && isReference(node, parent) && !declared.has(node.name) && !topNames.has(node.name)) globals.add(node.name);
  });
  return {
    name: fn.id.name,
    bytes: fn.end - fn.start,
    crossCalls: [...calls].filter((name) => topNames.has(name)).sort(),
    globals: [...globals].sort(),
    dom: /\b(?:document|window|HTMLElement|HTMLInputElement|HTMLButtonElement)\b/.test(text),
    firebase: /\b(?:firebase|firestore|\bdb\b|\bauth\b)/i.test(text),
    async: fn.async === true
  };
}

const groups = {
  snow: /snow|neve/i,
  weather: /weather|meteo|rain|radar/i,
  safety: /safety|sicurezza|alert/i,
  documents: /document|privateDocs|\bpos\b/i,
  notes: /note|segnal/i,
  hours: /hours|ore/i,
  calendar: /calendar/i,
  global: /global/i,
  map: /map/i
};

const report = {};
for (const [group, pattern] of Object.entries(groups)) {
  const items = topFns.filter((fn) => pattern.test(fn.id.name)).map(analyze);
  const names = new Set(items.map((item) => item.name));
  const externalCalls = [...new Set(items.flatMap((item) => item.crossCalls.filter((name) => !names.has(name))))];
  const globals = [...new Set(items.flatMap((item) => item.globals))];
  report[group] = {
    functions: items.length,
    bytes: items.reduce((sum, item) => sum + item.bytes, 0),
    domFunctions: items.filter((item) => item.dom).length,
    firebaseFunctions: items.filter((item) => item.firebase).length,
    asyncFunctions: items.filter((item) => item.async).length,
    externalCallCount: externalCalls.length,
    globalDependencyCount: globals.length,
    names: items.map((item) => item.name).sort(),
    externalCalls: externalCalls.sort()
  };
}

const ranked = Object.entries(report)
  .filter(([, item]) => item.functions >= 3)
  .map(([name, item]) => ({ name, score: item.externalCallCount + item.globalDependencyCount + item.domFunctions * 2 + item.firebaseFunctions * 3, ...item }))
  .sort((a, b) => a.score - b.score || a.bytes - b.bytes);

const output = { generatedAt: new Date().toISOString(), appBytes: source.length, topLevelFunctionCount: topFns.length, ranked, groups: report };
fs.mkdirSync(path.join(ROOT, "docs"), { recursive: true });
fs.writeFileSync(path.join(ROOT, "docs/phase33-next-module-audit.json"), JSON.stringify(output, null, 2) + "\n");
console.log(JSON.stringify(ranked.slice(0, 6), null, 2));

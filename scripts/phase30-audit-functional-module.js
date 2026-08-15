#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");

const ROOT = path.resolve(__dirname, "..");
const source = fs.readFileSync(path.join(ROOT, "app.js"), "utf8");
const ast = acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });

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
  const names = new Set([fn.id?.name].filter(Boolean));
  for (const p of fn.params || []) walk(p, (n) => { if (n.type === "Identifier") names.add(n.name); });
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

const topFns = ast.body.filter((node) => node.type === "FunctionDeclaration" && node.id?.name);
const topNames = new Set(topFns.map((fn) => fn.id.name));

function analyze(fn) {
  const declared = declaredNames(fn);
  const external = new Set();
  const calls = new Set();
  const text = source.slice(fn.start, fn.end);
  walk(fn.body, (node, parent) => {
    if (node.type === "CallExpression" && node.callee?.type === "Identifier") calls.add(node.callee.name);
    if (node.type === "Identifier" && isReferenceIdentifier(node, parent) && !declared.has(node.name)) {
      external.add(node.name);
    }
  });
  const topLevelCalls = [...calls].filter((name) => topNames.has(name));
  const globalDeps = [...external].filter((name) => !topNames.has(name));
  return {
    name: fn.id.name,
    bytes: fn.end - fn.start,
    topLevelCalls: topLevelCalls.sort(),
    globalDeps: globalDeps.sort(),
    touchesDom: /\b(?:document|window|HTMLElement|HTMLInputElement|HTMLButtonElement)\b/.test(text),
    touchesFirebase: /\b(?:firebase|firestore|auth)\b/i.test(text),
    async: fn.async === true
  };
}

function cluster(pattern) {
  const seed = topFns.filter((fn) => pattern.test(fn.id.name) || pattern.test(source.slice(fn.start, fn.end)));
  const seedNames = new Set(seed.map((fn) => fn.id.name));
  // Include first-level helpers called only by seed functions when their names also suggest the same feature.
  return seed.map(analyze).sort((a, b) => a.name.localeCompare(b.name));
}

const report = {
  generatedAt: new Date().toISOString(),
  appBytes: source.length,
  topLevelFunctionCount: topFns.length,
  clusters: {
    atex: cluster(/atex/i),
    worklimate: cluster(/worklimate/i)
  }
};

for (const [key, items] of Object.entries(report.clusters)) {
  const names = new Set(items.map((item) => item.name));
  report.clusters[`${key}Summary`] = {
    functions: items.length,
    bytes: items.reduce((sum, item) => sum + item.bytes, 0),
    crossClusterCalls: [...new Set(items.flatMap((item) => item.topLevelCalls.filter((name) => !names.has(name))))].sort(),
    globalDependencies: [...new Set(items.flatMap((item) => item.globalDeps))].sort(),
    domFunctions: items.filter((item) => item.touchesDom).length,
    firebaseFunctions: items.filter((item) => item.touchesFirebase).length,
    asyncFunctions: items.filter((item) => item.async).length
  };
}

fs.mkdirSync(path.join(ROOT, "docs"), { recursive: true });
fs.writeFileSync(path.join(ROOT, "docs", "phase30-module-audit.json"), JSON.stringify(report, null, 2) + "\n", "utf8");
console.log(JSON.stringify(report.clusters, null, 2));

#!/usr/bin/env node
"use strict";
const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");
const root = path.resolve(__dirname, "..");
const source = fs.readFileSync(path.join(root, "app.js"), "utf8");
const ast = acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
const functions = ast.body
  .filter((node) => node.type === "FunctionDeclaration" && node.id?.name && /snow|neve/i.test(node.id.name))
  .map((node) => ({ name: node.id.name, bytes: node.end - node.start, async: node.async === true, source: source.slice(node.start, node.end) }));
for (const item of functions) {
  item.dom = /\b(?:document|window|HTMLElement)\b/.test(item.source);
  item.firebase = /\b(?:firebase|\bdb\b|firestore|\bauth\b)/i.test(item.source);
  delete item.source;
}
console.log(JSON.stringify({ count: functions.length, bytes: functions.reduce((sum, item) => sum + item.bytes, 0), functions }, null, 2));

#!/usr/bin/env node
"use strict";
const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");
const filePath = path.resolve(__dirname, "..", "app.js");
const targets = ["canTriggerImpiantoWhatsApp", "setWhazzupPreparingFeedback"];
let source = fs.readFileSync(filePath, "utf8");
function collect(node, out = []) {
  if (!node || typeof node !== "object") return out;
  if (node.type === "FunctionDeclaration" && node.id?.name) out.push(node);
  for (const [key, value] of Object.entries(node)) {
    if (["start", "end", "loc"].includes(key)) continue;
    if (Array.isArray(value)) value.forEach((item) => collect(item, out));
    else if (value && typeof value === "object" && typeof value.type === "string") collect(value, out);
  }
  return out;
}
const ast = acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
const declarations = collect(ast);
const ranges = [];
for (const name of targets) {
  const matches = declarations.filter((node) => node.id.name === name);
  if (matches.length !== 1) throw new Error(`${name}: dichiarazioni=${matches.length}`);
  const refs = source.match(new RegExp(`\\b${name}\\b`, "g")) || [];
  if (refs.length !== 1) throw new Error(`${name}: riferimenti=${refs.length}; rimozione annullata`);
  let { start, end } = matches[0];
  while (end < source.length && /[ \t]/.test(source[end])) end += 1;
  if (source[end] === "\r") end += 1;
  if (source[end] === "\n") end += 1;
  if (source[end] === "\r") end += 1;
  if (source[end] === "\n") end += 1;
  ranges.push({ start, end, name });
}
ranges.sort((a, b) => b.start - a.start);
for (const range of ranges) {
  source = source.slice(0, range.start) + source.slice(range.end);
  console.log(`RIMOSSA ${range.name}`);
}
acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
fs.writeFileSync(filePath, source, "utf8");
console.log("Fase 25: rimossi gli ultimi 2 helper orfani a cascata.");

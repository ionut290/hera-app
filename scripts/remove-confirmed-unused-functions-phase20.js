#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const acorn = require("acorn");

const ROOT = path.resolve(__dirname, "..");
const TARGETS = {
  "active-commesse-first-boot-guard.js": ["deliverEmptySnapshot"],
  "app.js": [
    "buildBannedWhatsAppUrl",
    "getPlatformUserIdentityParts",
    "parseGoogleSheetId",
    "createWorklimateButton",
    "openGlobalSegnalazioneModal",
    "formatCompactImpiantoWeatherRiskLine",
    "formatImpiantoRainLine",
    "getImpiantoWeatherAlertLine",
    "getImpiantoWeatherLineClass",
    "renderSimpleList"
  ],
  "today-summary-interactions.js": ["getPlannedHours", "getAlertGroups"]
};

function collectFunctionDeclarations(node, out = []) {
  if (!node || typeof node !== "object") return out;
  if (node.type === "FunctionDeclaration" && node.id?.name) out.push(node);
  for (const [key, value] of Object.entries(node)) {
    if (key === "start" || key === "end" || key === "loc") continue;
    if (Array.isArray(value)) value.forEach((item) => collectFunctionDeclarations(item, out));
    else if (value && typeof value === "object" && typeof value.type === "string") collectFunctionDeclarations(value, out);
  }
  return out;
}

for (const [relativePath, names] of Object.entries(TARGETS)) {
  const filePath = path.join(ROOT, relativePath);
  let source = fs.readFileSync(filePath, "utf8");
  const ast = acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
  const declarations = collectFunctionDeclarations(ast);
  const ranges = [];

  for (const name of names) {
    const matches = declarations.filter((node) => node.id.name === name);
    if (matches.length !== 1) throw new Error(`${relativePath}: ${name} atteso una volta, trovato ${matches.length}`);
    const refCount = (source.match(new RegExp(`\\b${name}\\b`, "g")) || []).length;
    if (refCount !== 1) throw new Error(`${relativePath}: ${name} non è più un candidato sicuro, riferimenti=${refCount}`);
    let { start, end } = matches[0];
    while (end < source.length && (source[end] === " " || source[end] === "\t")) end += 1;
    if (source[end] === "\r") end += 1;
    if (source[end] === "\n") end += 1;
    if (source[end] === "\r") end += 1;
    if (source[end] === "\n") end += 1;
    ranges.push({ start, end, name });
  }

  ranges.sort((a, b) => b.start - a.start);
  for (const range of ranges) {
    source = source.slice(0, range.start) + source.slice(range.end);
    console.log(`RIMOSSA ${relativePath} :: ${range.name}`);
  }
  acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", allowHashBang: true });
  fs.writeFileSync(filePath, source, "utf8");
}

console.log("Fase 20: funzioni inutilizzate rimosse con parsing AST.");

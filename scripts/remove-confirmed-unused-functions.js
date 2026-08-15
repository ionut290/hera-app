#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const TARGETS = {
  "app.js": [
    "openBannedAccessRequest",
    "isCurrentUserBanned",
    "urlBase64ToUint8Array",
    "getCurrentUserIdentityParts",
    "getHoursOperatorForCurrentUser",
    "parseGoogleSheetId",
    "getCommesseErrorMessage",
    "getCommessaHoursTotal",
    "createWorklimateButton",
    "shareGlobalImpiantoViaWhatsapp",
    "toggleCommessaNoteForm",
    "formatCompactImpiantoWeatherRiskLine",
    "formatImpiantoRainLine",
    "getImpiantoWeatherAlertLine",
    "getImpiantoWeatherLineClass",
    "formatWeatherDetailValue",
    "formatWeatherAmount",
    "renderSimpleList",
    "getTimestampDate",
    "openWeatherModal",
    "renderSnowServiceList"
  ],
  "today-summary-interactions.js": [
    "getPlannedHours",
    "getAlertGroups",
    "getLiveWorkedMinutes"
  ],
  "active-commesse-first-boot-guard.js": [
    "deliverEmptySnapshot"
  ]
};

function findFunctionRange(source, name) {
  const pattern = new RegExp(`(^|\\n)([ \\t]*)function\\s+${name.replace(/[.*+?^${}()|[\\]\\\\]/g, "\\\\$&")}\\s*\\(`, "m");
  const match = pattern.exec(source);
  if (!match) return null;
  const start = match.index + (match[1] ? match[1].length : 0);
  const declarationStart = start + match[2].length;
  const openBrace = source.indexOf("{", declarationStart);
  if (openBrace < 0) throw new Error(`Graffa iniziale non trovata per ${name}`);

  let depth = 0;
  let mode = "code";
  let escaped = false;
  for (let i = openBrace; i < source.length; i += 1) {
    const ch = source[i];
    const next = source[i + 1];

    if (mode === "lineComment") {
      if (ch === "\n") mode = "code";
      continue;
    }
    if (mode === "blockComment") {
      if (ch === "*" && next === "/") { mode = "code"; i += 1; }
      continue;
    }
    if (mode === "single" || mode === "double" || mode === "template") {
      if (escaped) { escaped = false; continue; }
      if (ch === "\\") { escaped = true; continue; }
      if ((mode === "single" && ch === "'") || (mode === "double" && ch === '"') || (mode === "template" && ch === "`")) mode = "code";
      continue;
    }

    if (ch === "/" && next === "/") { mode = "lineComment"; i += 1; continue; }
    if (ch === "/" && next === "*") { mode = "blockComment"; i += 1; continue; }
    if (ch === "'") { mode = "single"; continue; }
    if (ch === '"') { mode = "double"; continue; }
    if (ch === "`") { mode = "template"; continue; }
    if (ch === "{") depth += 1;
    if (ch === "}") {
      depth -= 1;
      if (depth === 0) {
        let end = i + 1;
        while (source[end] === " " || source[end] === "\t") end += 1;
        if (source[end] === "\r") end += 1;
        if (source[end] === "\n") end += 1;
        if (source[end] === "\r") end += 1;
        if (source[end] === "\n") end += 1;
        return { start, end };
      }
    }
  }
  throw new Error(`Fine funzione non trovata per ${name}`);
}

let totalRemoved = 0;
for (const [relativePath, names] of Object.entries(TARGETS)) {
  const filePath = path.join(ROOT, relativePath);
  let source = fs.readFileSync(filePath, "utf8");
  for (const name of names) {
    const occurrences = source.match(new RegExp(`\\b${name}\\b`, "g")) || [];
    if (occurrences.length !== 1) throw new Error(`${name}: attesi 1 riferimento prima della rimozione, trovati ${occurrences.length}`);
    const range = findFunctionRange(source, name);
    if (!range) throw new Error(`${name}: dichiarazione non trovata`);
    source = source.slice(0, range.start) + source.slice(range.end);
    totalRemoved += 1;
    console.log(`RIMOSSA ${relativePath} :: ${name}`);
  }
  fs.writeFileSync(filePath, source);
}

console.log(`Rimosse ${totalRemoved} funzioni confermate come inutilizzate.`);

#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const SKIP_DIRS = new Set(["node_modules", ".git", "android", "dist", "build", ".netlify"]);
const SKIP_FILES = new Set([path.join("scripts", "audit-unused-functions.js")]);

function walk(dir, out = []) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name.startsWith(".") && entry.name !== ".github") continue;
    const full = path.join(dir, entry.name);
    const rel = path.relative(ROOT, full);
    if (entry.isDirectory()) {
      if (!SKIP_DIRS.has(entry.name)) walk(full, out);
      continue;
    }
    if (!entry.isFile() || !entry.name.endsWith(".js")) continue;
    if (SKIP_FILES.has(rel)) continue;
    out.push({ rel, text: fs.readFileSync(full, "utf8") });
  }
  return out;
}

const files = walk(ROOT);
const corpus = files.map((file) => file.text).join("\n");
const declarations = [];
const declarationPattern = /(?:^|[\n;{}]\s*)function\s+([A-Za-z_$][\w$]*)\s*\(/g;

for (const file of files) {
  let match;
  while ((match = declarationPattern.exec(file.text))) {
    const prefix = file.text.slice(0, match.index);
    const line = prefix.split("\n").length;
    declarations.push({ name: match[1], file: file.rel, line });
  }
}

const candidates = declarations
  .map((decl) => {
    const escaped = decl.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const refs = corpus.match(new RegExp(`\\b${escaped}\\b`, "g")) || [];
    return { ...decl, refs: refs.length };
  })
  .filter((item) => item.refs === 1)
  .sort((a, b) => a.file.localeCompare(b.file) || a.line - b.line);

console.log(`Scansionati ${files.length} file JavaScript e ${declarations.length} dichiarazioni function.`);
if (!candidates.length) {
  console.log("Nessuna funzione dichiarata con un solo riferimento testuale.");
  process.exit(0);
}
console.log(`Candidati a funzione mai usata: ${candidates.length}`);
for (const item of candidates) console.log(`UNUSED_CANDIDATE ${item.file}:${item.line} ${item.name}`);

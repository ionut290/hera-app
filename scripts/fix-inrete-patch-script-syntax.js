#!/usr/bin/env node
"use strict";

const fs = require("node:fs");

const path = "scripts/apply-inrete-modena-canonical-count-fix.js";
let source = fs.readFileSync(path, "utf8");

const replacements = [
  [
    'if (sap) return `SAP:${sap}`;',
    'if (sap) return "SAP:" + sap;'
  ],
  [
    'if (direct) return `ID:${direct}`;',
    'if (direct) return "ID:" + direct;'
  ],
  [
    'const anagrafica = normalize(`${row?.denominazione ?? row?.nome ?? ""}|${row?.comune ?? ""}|${row?.indirizzo ?? row?.via ?? ""}`);',
    'const anagrafica = normalize(String(row?.denominazione ?? row?.nome ?? "") + "|" + String(row?.comune ?? "") + "|" + String(row?.indirizzo ?? row?.via ?? ""));'
  ],
  [
    'if (anagrafica) return `ANAG:${anagrafica}`;',
    'if (anagrafica) return "ANAG:" + anagrafica;'
  ],
  [
    'return `ROW:${String(row?.id ?? row?.numeroProgressivoRiga ?? "").trim()}`;',
    'return "ROW:" + String(row?.id ?? row?.numeroProgressivoRiga ?? "").trim();'
  ],
  [
    'id: `work-${index + 1}`,',
    'id: "work-" + (index + 1),'
  ]
];

for (const [before, after] of replacements) {
  if (source.includes(after)) continue;
  const count = source.split(before).length - 1;
  if (count !== 1) {
    throw new Error(`Sostituzione non sicura: attesa 1 occorrenza di ${before}, trovate ${count}`);
  }
  source = source.replace(before, after);
}

fs.writeFileSync(path, source, "utf8");
console.log("Sintassi dell’applicatore Modena corretta.");

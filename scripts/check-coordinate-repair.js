#!/usr/bin/env node
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const repairSource = fs.readFileSync(path.join(__dirname, "..", "coordinate-repair.js"), "utf8");
const accountingSource = fs.readFileSync(path.join(__dirname, "..", "accounting-v2.js"), "utf8");
const appSource = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
const cssSource = fs.readFileSync(path.join(__dirname, "..", "accounting-v2.css"), "utf8");
const context = { window: {} };
vm.createContext(context);
vm.runInContext(repairSource, context);
const tools = context.window.HeraCoordinateRepair;

const cases = [
  ["coordinate separate con virgola", tools.diagnose("44,6339", "11,6679"), { valid: true, repaired: false, latitude: 44.6339, longitude: 11.6679 }],
  ["coordinate nello stesso campo", tools.diagnose("44.6339, 11.6679", ""), { valid: true, repaired: true, latitude: 44.6339, longitude: 11.6679 }],
  ["coordinate italiane invertite", tools.diagnose("11.6679", "44.6339"), { valid: true, repaired: true, latitude: 44.6339, longitude: 11.6679 }],
  ["coordinate non valide conservate", tools.diagnose("errore", "11.6679"), { valid: false, repaired: false, rawLatitude: "errore", rawLongitude: "11.6679" }],
  ["coordinate mancanti segnalate", tools.diagnose("", ""), { valid: false, status: "MISSING" }]
];

let failed = false;
for (const [name, actual, expected] of cases) {
  const passed = Object.entries(expected).every(([key, value]) => actual[key] === value);
  console.log(`${passed ? "OK" : "FAIL"} ${name}`);
  failed ||= !passed;
}

const checks = [
  ["import conserva valori originali", accountingSource.includes("coordinateLatitudineOriginale") && accountingSource.includes("coordinateLongitudineOriginale")],
  ["import segnala coordinate da verificare", accountingSource.includes("Coordinate da verificare")],
  ["righe problematiche evidenziate", accountingSource.includes("coordinate-problem-row") && cssSource.includes(".coordinate-problem-row>td")],
  ["messaggio problema rosso", accountingSource.includes("coordinate-problem-copy") && cssSource.includes("color:#dc2626")],
  ["impianti esistenti diagnosticati", accountingSource.includes("state.work.map(joined).filter(hasCoordinateProblem)")],
  ["coordinate unite supportate", accountingSource.includes("coordinateUnica")],
  ["import principale conserva coordinate originali", appSource.includes("coordinateLatitudineOriginale: coordinateDiagnosis.rawLatitude")],
  ["import principale aggiorna coordinate esistenti", appSource.includes("const coordinateChanged = hasImportedCoordinates")]
];

for (const [name, passed] of checks) {
  console.log(`${passed ? "OK" : "FAIL"} ${name}`);
  failed ||= !passed;
}

if (failed) process.exit(1);

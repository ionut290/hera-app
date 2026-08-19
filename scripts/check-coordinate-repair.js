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
const runtime = context.window.HeraCoordinateReadRuntime;

const cases = [
  ["coordinate separate con virgola", tools.diagnose("44,6339", "11,6679"), { valid: true, repaired: false, latitude: 44.6339, longitude: 11.6679 }],
  ["coordinate nello stesso campo", tools.diagnose("44.6339, 11.6679", ""), { valid: true, repaired: true, latitude: 44.6339, longitude: 11.6679 }],
  ["coordinate italiane invertite", tools.diagnose("11.6679", "44.6339"), { valid: true, repaired: true, latitude: 44.6339, longitude: 11.6679 }],
  ["coordinate non valide conservate", tools.diagnose("errore", "11.6679"), { valid: false, repaired: false, rawLatitude: "errore", rawLongitude: "11.6679" }],
  ["coordinate mancanti segnalate", tools.diagnose("", ""), { valid: false, status: "MISSING" }]
];

const recordCases = [
  ["coppia INRETE nel campo latitudine originale", tools.resolveRecord({ coordinateLatitudineOriginale: "10,9252 / 44,6471" }), { valid: true, repaired: true, latitude: 44.6471, longitude: 10.9252 }],
  ["coppia INRETE nel campo GPS Y", tools.resolveRecord({ gpsY: "10,9252 / 44,6471" }), { valid: true, repaired: true, latitude: 44.6471, longitude: 10.9252 }],
  ["coppia INRETE nel campo GPS X", tools.resolveRecord({ gpsX: "10,9252 / 44,6471" }), { valid: true, repaired: true, latitude: 44.6471, longitude: 10.9252 }],
  ["coordinate INRETE nei campi extra storici", tools.resolveRecord({ extraFields: { coordinategpsy: "44,5421", coordinategpsx: "10,8112" } }), { valid: true, latitude: 44.5421, longitude: 10.8112, source: "extraFields.coordinategpsy + coordinategpsx" }]
];

const physicalRows = [
  { id: "plant-formigine", denominazione: "REMI FORMIGINE", comune: "Formigine", indirizzo: "Via per Sassuolo", latitudine: 44.55, longitudine: 10.85 },
  { id: "plant-castellaro", denominazione: "REMI CASTELLARO", comune: "Sala Bolognese", indirizzo: "Via Stelloni", latitudine: 44.63, longitudine: 11.25 },
  { id: "plant-braida", denominazione: "REMI BRAIDA", comune: "Sassuolo", indirizzo: "Via Braida", latitudine: 44.54, longitudine: 10.78 }
];

const physicalMatchCases = [
  ["collegamento da migrationSourceId", runtime.matchPhysicalRecord({ migrationSourceId: "plant-formigine::A11" }, physicalRows)?.id, "plant-formigine"],
  ["collegamento da id lavorazione", runtime.matchPhysicalRecord({ id: "plant-castellaro__A11" }, physicalRows)?.id, "plant-castellaro"],
  ["collegamento da nome univoco senza comune e indirizzo", runtime.matchPhysicalRecord({ denominazione: "Remi  Braida" }, physicalRows)?.id, "plant-braida"],
  ["nome duplicato non collegato", runtime.matchPhysicalRecord({ denominazione: "REMI BRAIDA" }, [...physicalRows, { id: "plant-braida-2", denominazione: "REMI BRAIDA" }]), null]
];

const historicalBatch = runtime.normalizeRows(Array.from({ length: 30 }, (_, index) => ({
  id: `remi-${index + 1}`,
  extraFields: {
    "Coordinate GPS(Y)": String(44.50 + (index / 1000)).replace(".", ","),
    "Coordinate GPS(X)": String(10.70 + (index / 1000)).replace(".", ",")
  }
})));

let failed = false;
for (const [name, actual, expected] of cases) {
  const passed = Object.entries(expected).every(([key, value]) => actual[key] === value);
  console.log(`${passed ? "OK" : "FAIL"} ${name}`);
  failed ||= !passed;
}

for (const [name, actual, expected] of recordCases) {
  const passed = Object.entries(expected).every(([key, value]) => actual[key] === value);
  console.log(`${passed ? "OK" : "FAIL"} ${name}`);
  failed ||= !passed;
}

for (const [name, actual, expected] of physicalMatchCases) {
  const passed = actual === expected;
  console.log(`${passed ? "OK" : "FAIL"} ${name}`);
  failed ||= !passed;
}

const checks = [
  ["tutte le 30 righe storiche INRETE vengono normalizzate", historicalBatch.length === 30 && historicalBatch.every((row) => tools.resolveRecord(row).valid)],
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

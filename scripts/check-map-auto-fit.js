#!/usr/bin/env node
const fs = require("fs");
const path = require("path");

const source = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");

function extractFunction(name) {
  const start = source.indexOf("function " + name + "(");
  if (start < 0) return "";
  const bodyStart = source.indexOf("{", start);
  let depth = 0;
  for (let index = bodyStart; index < source.length; index += 1) {
    if (source[index] === "{") depth += 1;
    if (source[index] === "}") depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }
  return "";
}

const parserSource = extractFunction("parseImpiantoMapCoordinate");
const coordinatesSource = extractFunction("getImpiantoMapCoordinates");
if (!parserSource || !coordinatesSource) {
  console.error("FAIL Funzioni di validazione coordinate mappa mancanti");
  process.exit(1);
}

const getCoordinates = new Function(
  parserSource + "\n" + coordinatesSource + "\nreturn getImpiantoMapCoordinates;"
)();

const tests = [
  ["coordinate decimali valide", { gpsY: "44.6339", gpsX: "11.6679" }, [44.6339, 11.6679]],
  ["virgola decimale supportata", { gpsY: "44,6339", gpsX: "11,6679" }, [44.6339, 11.6679]],
  ["campi vuoti esclusi", { gpsY: "", gpsX: "" }, null],
  ["latitudine zero esclusa", { gpsY: 0, gpsX: 11.3426 }, null],
  ["longitudine zero esclusa", { gpsY: 44.4949, gpsX: 0 }, null],
  ["coordinate fuori intervallo escluse", { gpsY: 120, gpsX: 11.3426 }, null]
];

let failed = false;
for (const [name, input, expected] of tests) {
  const actual = getCoordinates(input);
  const passed = JSON.stringify(actual) === JSON.stringify(expected);
  console.log((passed ? "OK " : "FAIL ") + name);
  failed ||= !passed;
}

const renderMap = extractFunction("renderMap");
const markerFactory = extractFunction("addImpiantoMarkerToMapLayer");
const guard = fs.readFileSync(path.join(__dirname, "..", "accounting-view-guard.js"), "utf8");
const structuralChecks = [
  ["fitBounds usa coordinate validate", renderMap.includes("if (coordinates) impiantiBounds.push(coordinates)")],
  ["firma zoom usa impianti e coordinate", renderMap.includes("coordinates?.[0]")],
  ["marker rifiuta coordinate non valide", markerFactory.includes("if (!coordinates) return null")],
  ["guard esterno mappa rimosso", !guard.includes("installMapAutoFitGuard")]
];

for (const [name, passed] of structuralChecks) {
  console.log((passed ? "OK " : "FAIL ") + name);
  failed ||= !passed;
}

if (failed) process.exit(1);

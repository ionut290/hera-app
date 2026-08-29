#!/usr/bin/env node

const fs = require("fs");
const path = require("path");

const rootDir = path.resolve(__dirname, "..");
const appSource = fs.readFileSync(path.join(rootDir, "app.js"), "utf8");
const start = appSource.indexOf("function renderGlobalMap() {");
const end = appSource.indexOf("\nfunction ", start + 1);

if (start < 0 || end < 0) {
  throw new Error("Funzione renderGlobalMap non trovata.");
}

const renderGlobalMapSource = appSource.slice(start, end);
const requiredSnippets = [
  "const coordinates = [Number(impianto.gpsY), Number(impianto.gpsX)];",
  "if (!isValidLatLon(coordinates[0], coordinates[1])) return;",
  "L.marker(coordinates,",
  "bounds.push(coordinates);"
];

requiredSnippets.forEach((snippet) => {
  if (!renderGlobalMapSource.includes(snippet)) {
    throw new Error(`Protezione coordinate mappa Global mancante: ${snippet}`);
  }
});

if (renderGlobalMapSource.includes("bounds.push([impianto.gpsY, impianto.gpsX])")) {
  throw new Error("La mappa Global usa ancora coordinate non normalizzate nei bounds.");
}

console.log("Controllo coordinate mappa Global superato.");

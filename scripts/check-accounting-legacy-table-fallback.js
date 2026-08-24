#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const core = require("../inrete-work-items-v2.js");

const source = fs.readFileSync("accounting-v2.js", "utf8");

assert.match(
  source,
  /const legacy=state\.operationalPlants\.length\?state\.operationalPlants:cachedLegacy;/,
  "Quando lavorazioni è vuota, la Gestione deve usare prima gli impianti operativi reali"
);
assert.match(source, /const physicalById=new Map\(state\.plants\.map/);
assert.match(source, /state\.work=state\.plants\.flatMap\(p=>core\.adaptLegacyPlantToWorkItems/);

const blockMatch = source.match(/if\(!state\.work\.length\)\{[\s\S]*?\n    \}\n    \/\/ Risoluzione Firebase\/JavaScript/);
assert.ok(blockMatch, "Blocco fallback Gestione non trovato");
assert.doesNotMatch(blockMatch[0], /\.set\s*\(/, "Il fallback non deve scrivere su Firestore");
assert.doesNotMatch(blockMatch[0], /\.delete\s*\(/, "Il fallback non deve cancellare dati Firestore");

const operationalPlants = Array.from({ length: 94 }, (_, index) => ({
  id: `ferrara_${index + 1}`,
  physicalPlantId: `ferrara_${index + 1}`,
  numeroProgressivo: index + 1,
  idSap: String(100000 + index),
  denominazione: `IMPIANTO FERRARA ${index + 1}`,
  comune: "FERRARA",
  indirizzo: `Via ${index + 1}`,
  gpsY: 44.83 + index / 10000,
  gpsX: 11.62 + index / 10000,
  codicePrezzo: "A.01",
  tipologiaIntervento: "Sfalcio",
  stato: "FATTO",
  statoGenerale: "FATTO",
  done: true
}));

const physicalPlants = operationalPlants.slice(0, 30).map((plant) => ({
  id: plant.id,
  idSap: plant.idSap,
  denominazione: plant.denominazione,
  comune: plant.comune,
  indirizzo: plant.indirizzo
}));
const cachedLegacy = [];
const legacy = operationalPlants.length ? operationalPlants : cachedLegacy;
const physicalById = new Map(physicalPlants.map((plant) => [String(plant.id || ""), plant]));
const plants = legacy.map((plant, index) => {
  const linkedId = String(plant.physicalPlantId || plant.id || "").trim();
  const physical = physicalById.get(linkedId) || {};
  const merged = { ...physical, ...plant };
  const id = linkedId || String(physical.id || `legacy_${index + 1}`);
  return {
    ...merged,
    id,
    legacy: true,
    numeroProgressivoImpianto: merged.numeroProgressivoImpianto ?? merged.numeroProgressivo ?? index + 1,
    latitudine: merged.coordinateLatitudineOriginale ?? merged.latitudine ?? merged.gpsY,
    longitudine: merged.coordinateLongitudineOriginale ?? merged.longitudine ?? merged.gpsX
  };
});
const work = plants.flatMap((plant) => core.adaptLegacyPlantToWorkItems({
  ...plant,
  numeroProgressivoRiga: plant.numeroProgressivoRiga ?? plant.numeroProgressivo ?? plant.numeroProgressivoImpianto,
  frequenzaAnnua: plant.frequenzaAnnua || "",
  note: plant.note || plant.noteImpianto || ""
}));

assert.equal(plants.length, 94, "Ferrara deve mantenere tutti i 94 impianti reali");
assert.equal(work.length, 94, "Con una lavorazione legacy per impianto la tabella deve mostrare 94 righe");
assert.equal(work.filter((row) => row.stato === "FATTO").length, 94, "Gli stati FATTO devono essere conservati");
assert.equal(work[93].impiantoId, "ferrara_94");
assert.equal(work[93].latitudine, operationalPlants[93].gpsY);
assert.equal(work[93].longitudine, operationalPlants[93].gpsX);

console.log("✅ Gestione contabilità: 94 impianti operativi diventano 94 righe visibili senza scritture Firestore.");

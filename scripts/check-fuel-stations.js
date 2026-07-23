const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const fuel = require("../fuel-stations-core.js");

const variants = {
  cng: ["Metano", " CNG ", "GNC", "Mètano + benzina"],
  lpg: ["GPL", "lpg", " GPL + benzina "],
  diesel: ["Diesel", "gasolio"],
  petrol: ["Benzina"],
  electric: ["Elettrico"]
};
for (const [expected, values] of Object.entries(variants)) values.forEach((value) => assert.equal(fuel.normalizeFuel(value), expected));

const elements = [
  { type: "node", id: 1, lat: 45, lon: 9, tags: { amenity: "fuel", brand: "Q8", name: "Q8 CNG", "fuel:cng": "yes", "addr:street": "Via Roma", "addr:housenumber": "1" } },
  { type: "node", id: 2, lat: 45.1, lon: 9, tags: { amenity: "fuel", brand: "Eni", "fuel:lpg": "yes" } },
  { type: "node", id: 3, lat: 45.2, lon: 9, tags: { amenity: "fuel", brand: "Q8", "fuel:diesel": "yes" } },
  { type: "node", id: 4, lat: 45.3, lon: 9, tags: { amenity: "fuel", brand: "Agip", "fuel:octane_95": "yes" } },
  { type: "node", id: 5, lat: 45.4, lon: 9, tags: { amenity: "fuel", brand: "Q8" } }
];
const distance = (_a, _b, lat) => lat;
assert.deepEqual(fuel.parseStations(elements, "cng", { lat: 0, lng: 0 }, distance).map((x) => x.id), ["node-1"]);
assert.deepEqual(fuel.parseStations(elements, "lpg", { lat: 0, lng: 0 }, distance).map((x) => x.id), ["node-2"]);
assert.deepEqual(fuel.parseStations(elements, "diesel", { lat: 0, lng: 0 }, distance).map((x) => x.id), ["node-3"]);
assert.deepEqual(fuel.parseStations(elements, "petrol", { lat: 0, lng: 0 }, distance).map((x) => x.id), ["node-4"]);
assert.equal(fuel.parseStations(elements, "cng", { lat: 0, lng: 0 }, distance).some((x) => x.id === "node-5"), false);
assert.match(fuel.buildQuery(45, 9, 5, "cng"), /around:5000,45,9/);
assert.match(fuel.buildQuery(45, 9, 30, "cng"), /around:30000,45,9/);
assert.match(fuel.buildQuery(45, 9, 50, "cng"), /around:50000,45,9/);
assert.match(fuel.buildQuery(45, 9, 30, "cng"), /Q8\|ENI\|Agip/);
assert.equal(fuel.formatAddress(elements[0].tags), "Via Roma 1");
const appSource = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
const indexSource = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
assert.match(appSource, /Posizione non disponibile\. Attiva la localizzazione e riprova\./);
assert.match(appSource, /Connessione assente\. Controlla Internet e riprova\./);
assert.match(appSource, /Nessun distributore Q8\/ENI che vende \$\{fuelLabel\} trovato nel raggio di \$\{radiusKm\} km\./);
assert.match(appSource, /createButton\("Riprova", \(\) => loadNearbyFuelStations\(\)\)/);
assert.match(appSource, /if \(fuelStationsLoadPromise\) return fuelStationsLoadPromise/);
assert.match(appSource, /if \(ui\.fuelRadius\) ui\.fuelRadius\.value = "5"/);
assert.match(appSource, /filter\(\(station\) => station\.distance <= radiusKm\)/);
assert.match(appSource, /class="mezzo-chip-btn squadra-conflict-name"/);
assert.match(indexSource, /<option value="5" selected>5 km \(predefinito\)<\/option>/);
assert.match(indexSource, /id="fuel-search-btn"/);
assert.match(appSource, /google\.com\/maps\/dir\/\?api=1&destination=/);
console.log("Fuel station normalization, strict filtering, radii and address checks passed.");

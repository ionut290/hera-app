const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const fuel = require("../fuel-stations-core.js");
const nationalCache = require("../fuel-stations-national-cache.js");
const nationalServerCache = require("../functions/fuel-stations-cache.js");
const fuelSearch = require("../fuel-stations-search.js");

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
const mimitResults = [
  {
    id: 101,
    name: "Q8 Metano",
    brand: "Q8",
    address: "Via Milano 2",
    location: { lat: 45.01, lng: 9.01 },
    insertDate: "2026-07-23T08:00:00+02:00",
    fuels: [
      { name: "Metano", price: 1.499, isSelf: false },
      { name: "Benzina", price: 1.899, isSelf: true }
    ]
  },
  {
    id: 102,
    name: "Eni GPL",
    brand: "Eni",
    address: "Via Torino 3",
    location: { lat: 45.02, lng: 9.02 },
    fuels: [{ name: "GPL", price: 0.749, isSelf: true }]
  },
  {
    id: 103,
    name: "Pompa bianca",
    brand: "Indipendente",
    address: "Via Napoli 4",
    location: { lat: 45.03, lng: 9.03 },
    fuels: [{ name: "Metano", price: 1.399, isSelf: true }]
  },
  {
    id: 104,
    name: "Q8 GNL",
    brand: "Q8",
    address: "Via Venezia 5",
    location: { lat: 45.04, lng: 9.04 },
    fuels: [{ name: "GNL", price: 1.299, isSelf: true }]
  }
];
assert.equal(fuel.MIMIT_API_URL, "https://carburanti.mise.gov.it/ospzApi/search/zone");
assert.deepEqual(fuel.buildMimitRequest(45, 9, 5), { points: [{ lat: 45, lng: 9 }], radius: 5 });
assert.equal(fuel.parseMimitStations(mimitResults, "cng", { lat: 0, lng: 0 }, distance).length, 1);
assert.match(fuel.parseMimitStations(mimitResults, "cng", { lat: 0, lng: 0 }, distance)[0].availableFuel, /metano • € 1\.499\/kg • servito/);
assert.deepEqual(fuel.parseMimitStations(mimitResults, "lpg", { lat: 0, lng: 0 }, distance).map((x) => x.id), ["mimit-102"]);
assert.equal(fuel.parseMimitStations(mimitResults, "cng", { lat: 0, lng: 0 }, distance).some((x) => x.id === "mimit-103"), false);
assert.equal(fuel.parseMimitStations(mimitResults, "cng", { lat: 0, lng: 0 }, distance).some((x) => x.id === "mimit-104"), false);

const nationalAnagrafica = [
  "Estrazione del 2026-07-23",
  "idimpianto|Gestore|Bandiera|Tipo Impianto|Nome Impianto|Indirizzo|Comune|Provincia|Latitudine|Longitudine",
  "201|Gestore Uno|Q8|Stradale|Q8 Roma|Via Roma 1|Roma|RM|41.900|12.500",
  "202|Gestore Due|Eni|Stradale|Eni Milano|Via Milano 2|Milano|MI|45.464|9.190",
  "203|Gestore Tre|Indipendente|Stradale|Pompa bianca|Via Napoli 3|Napoli|NA|40.851|14.268",
  "204|Gestore Quattro|Q8|Stradale|Coordinate errate|Via Zero 4|Roma|RM|0|0"
].join("\n");
const nationalPrices = [
  "Estrazione del 2026-07-23",
  "idimpianto|descCarburante|prezzo|isSelf|dtComu",
  "201|Metano|1.499|0|23/07/2026 08:00:00",
  "201|Benzina|1.899|1|23/07/2026 08:00:00",
  "202|GPL|0.749|1|23/07/2026 08:10:00",
  "203|Metano|1.399|1|23/07/2026 08:20:00",
  "204|Gasolio|1.799|1|23/07/2026 08:30:00"
].join("\n");
const nationalSnapshot = nationalServerCache.buildNationalSnapshot(nationalAnagrafica, nationalPrices, 123456);
assert.equal(nationalSnapshot.updatedAt, 123456);
assert.equal(nationalSnapshot.extractionDate, "2026-07-23");
assert.deepEqual(nationalSnapshot.stations.map((station) => station.id), ["201", "202"]);
assert.deepEqual(nationalSnapshot.stations[0].fuels.map((item) => item.name), ["Metano", "Benzina"]);
assert.equal(fuel.parseMimitStations(nationalSnapshot.stations, "cng", { lat: 0, lng: 0 }, distance).length, 1);
assert.equal(nationalServerCache.parseDelimitedTable("idimpianto;descCarburante;prezzo\n1;Gasolio;1,799").rows[0].prezzo, "1,799");
assert.equal(nationalCache.ENDPOINT, "/api/fuel-stations-italy");
assert.equal(nationalCache.FIREBASE_FUNCTION_REGION, "us-central1");
assert.deepEqual(nationalCache.endpointCandidates(), ["/api/fuel-stations-italy"]);
assert.equal(nationalCache.CACHE_TTL_MS, 24 * 60 * 60 * 1000);
assert.equal(nationalServerCache.CACHE_TTL_MS, 24 * 60 * 60 * 1000);
assert.equal(nationalCache.MAX_STALE_MS, 7 * 24 * 60 * 60 * 1000);
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
const swSource = fs.readFileSync(path.join(__dirname, "..", "sw.js"), "utf8");
const cacheSource = fs.readFileSync(path.join(__dirname, "..", "fuel-stations-national-cache.js"), "utf8");
const searchSource = fs.readFileSync(path.join(__dirname, "..", "fuel-stations-search.js"), "utf8");
const integrationSource = fs.readFileSync(path.join(__dirname, "..", "fuel-stations-integration.js"), "utf8");
const netlifyFunctionSource = fs.readFileSync(path.join(__dirname, "..", "netlify", "functions", "fuel-stations.js"), "utf8");
const netlifyConfigSource = fs.readFileSync(path.join(__dirname, "..", "netlify.toml"), "utf8");
const serverCacheSource = fs.readFileSync(path.join(__dirname, "..", "functions", "fuel-stations-cache.js"), "utf8");
const functionsSource = fs.readFileSync(path.join(__dirname, "..", "functions", "index.js"), "utf8");
const firebaseSource = fs.readFileSync(path.join(__dirname, "..", "firebase.json"), "utf8");
assert.match(appSource, /Posizione non disponibile\. Attiva la localizzazione e riprova\./);
assert.match(appSource, /Connessione assente\. Controlla Internet e riprova\./);
assert.match(appSource, /Nessun distributore Q8\/ENI che vende \$\{fuelLabel\} trovato nel raggio di \$\{radiusKm\} km\./);
assert.match(appSource, /createButton\("Riprova", \(\) => loadNearbyFuelStations\(\)\)/);
assert.match(appSource, /if \(fuelStationsLoadPromise\) return fuelStationsLoadPromise/);
assert.match(appSource, /if \(ui\.fuelRadius\) ui\.fuelRadius\.value = "5"/);
assert.match(appSource, /filter\(\(station\) => station\.distance <= radiusKm\)/);
assert.match(appSource, /fetchFuelStationsFromMimit\(position\.lat, position\.lng, radiusKm\)/);
assert.match(appSource, /uso la riserva OpenStreetMap/);
assert.match(appSource, /nationalCache\.findNearby\(fuel, position, radiusKm, haversine\)/);
assert.ok(appSource.indexOf("nationalCache.findNearby") < appSource.indexOf("fetchFuelStationsFromMimit(position.lat"));
assert.match(appSource, /source: "Archivio MIMIT salvato"/);
assert.match(indexSource, /archivio nazionale MIMIT salvato/);
assert.match(indexSource, /fuel-stations-national-cache\.js\?v=20260723/);
assert.match(indexSource, /fuel-stations-search\.js\?v=20260723a/);
assert.match(indexSource, /fuel-stations-integration\.js\?v=20260723a/);
assert.match(cacheSource, /ENDPOINT = "\/api\/fuel-stations-italy"/);
assert.match(cacheSource, /cloudfunctions\.net\/getFuelStationsItaly/);
assert.match(cacheSource, /content-type/);
assert.match(cacheSource, /requestIdleCallback/);
assert.match(serverCacheSource, /anagrafica_impianti_attivi\.csv/);
assert.match(serverCacheSource, /prezzo_alle_8\.csv/);
assert.match(functionsSource, /exports\.getFuelStationsItaly/);
assert.match(functionsSource, /exports\.refreshFuelStationsItaly/);
assert.match(functionsSource, /pubsub\.schedule\("30 3 \* \* \*"\)/);
assert.match(firebaseSource, /"source": "\/api\/fuel-stations-italy"/);
assert.match(firebaseSource, /"function": "getFuelStationsItaly"/);
assert.match(swSource, /hera-app-shell-v21/);
assert.match(swSource, /fuel-stations-national-cache\.js/);
assert.match(swSource, /fuel-stations-search\.js/);
assert.match(swSource, /fuel-stations-integration\.js/);
assert.deepEqual(fuelSearch.SAME_ORIGIN_ENDPOINTS, ["/api/fuel-stations/search", "/.netlify/functions/fuel-stations"]);
assert.equal(fuelSearch.OVERPASS_ENDPOINTS.length, 3);
assert.equal(fuelSearch.uniqueStations([
  { lat: 44.5, lon: 11.3, brandLabel: "Q8" },
  { lat: 44.5, lon: 11.3, brandLabel: "Q8" },
  { lat: 44.5, lon: 11.3, brandLabel: "ENI" }
]).length, 2);
assert.match(searchSource, /searchSameOrigin/);
assert.match(searchSource, /searchNationalCache/);
assert.match(searchSource, /searchMimit/);
assert.match(searchSource, /searchOverpass/);
assert.match(integrationSource, /HeraFuelStationSearch\.search/);
assert.match(integrationSource, /Ricerca tramite \$\{escapeHTML\(source\)\}/);
assert.match(netlifyFunctionSource, /exports\.handler/);
assert.match(netlifyFunctionSource, /fromMimit/);
assert.match(netlifyFunctionSource, /fromOverpass/);
assert.match(netlifyConfigSource, /from = "\/api\/fuel-stations\/search"/);
assert.match(appSource, /class="mezzo-chip-btn squadra-conflict-name"/);
assert.match(indexSource, /<option value="5" selected>5 km \(predefinito\)<\/option>/);
assert.match(indexSource, /id="fuel-search-btn"/);
assert.match(appSource, /google\.com\/maps\/dir\/\?api=1&destination=/);
console.log("Fuel station normalization, national MIMIT cache, strict filtering, radii and fallback checks passed.");

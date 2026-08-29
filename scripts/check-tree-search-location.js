#!/usr/bin/env node
const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const source = fs.readFileSync(path.join(root, "tree-search.js"), "utf8");
const css = fs.readFileSync(path.join(root, "tree-search.css"), "utf8");
const serviceWorker = fs.readFileSync(path.join(root, "sw.js"), "utf8");

const checks = [
  ["pulsante posizione presente", html.includes('id="tree-map-location-btn"')],
  ["pulsante collegato alla mappa", html.includes('id="tree-map-location-btn"') && html.includes('aria-controls="tree-map"')],
  ["richiesta GPS singola", source.includes("navigator.geolocation.getCurrentPosition") && !source.includes("navigator.geolocation.watchPosition")],
  ["alta precisione richiesta", source.includes("enableHighAccuracy: true")],
  ["marker posizione personale", source.includes("userLocationMarker = L.circleMarker")],
  ["cerchio precisione GPS", source.includes("userAccuracyCircle = L.circle")],
  ["centratura posizione", source.includes("map.setView(point, Math.max(map.getZoom(), 18)")],
  ["gestione permesso negato", source.includes("Permesso posizione negato")],
  ["listener pulsante presente", source.includes('mapLocationButton?.addEventListener("click", centerOnUserLocation)')],
  ["controllo adatto allo schermo intero", css.includes(".tree-map-card--fullscreen .tree-map-location-btn")],
  ["asset CSS aggiornato", html.includes("tree-search.css?v=20260829-location1") && serviceWorker.includes("tree-search.css?v=20260829-location1")],
  ["asset JS aggiornato", html.includes("tree-search.js?v=20260829-location1") && serviceWorker.includes("tree-search.js?v=20260829-location1")],
  ["nessuna operazione Firestore", !/firestore|\.collection\(|\.doc\(/i.test(source)]
];

let failed = false;
for (const [name, passed] of checks) {
  console.log(`${passed ? "OK" : "FAIL"} ${name}`);
  failed ||= !passed;
}

if (failed) process.exit(1);

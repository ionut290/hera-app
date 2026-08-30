#!/usr/bin/env node
const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const source = fs.readFileSync(path.join(root, "tree-search.js"), "utf8");
const css = fs.readFileSync(path.join(root, "tree-search.css"), "utf8");
const streetView = fs.readFileSync(path.join(root, "street-view-cards.js"), "utf8");
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
  ["solo 6 dettagli iniziali", source.includes('index >= 6 ? "tree-detail-extra"') && source.includes('index >= 6 ? " hidden"')],
  ["pulsante espansione dettagli", source.includes('class="btn tree-details-toggle"') && source.includes("MOSTRA SOLO I PRIMI 6 DETTAGLI")],
  ["pulsante invio Whazzup", source.includes('class="btn tree-whazzup-share"') && source.includes("buildTreeWhazzupMessage")],
  ["messaggio limitato ai primi 6 dettagli", source.includes("entries.slice(0, 6)")],
  ["navigazione inclusa nel messaggio", source.includes("📍 *NAVIGA VERSO L’ALBERO*") && source.includes("navigationUrl")],
  ["solo app WhatsApp installata", source.includes("whatsapp://send?text=") && !/wa\.me|web\.whatsapp\.com|api\.whatsapp\.com/.test(source)],
  ["plugin Android esistente riutilizzato", source.includes("Plugins?.HeraWhatsApp") && source.includes('plugin.open({ url: appUrl })')],
  ["pulsante vista 360 e percorso", source.includes('class="btn tree-street-view"') && source.includes("openTreeStreetView")],
  ["coordinate albero passate alla vista 360", source.includes("openForCoordinates") && source.includes("lng: Number(point.lon)")],
  ["API Street View riutilizzabile", streetView.includes("function openForCoordinates") && streetView.includes("openForCoordinates };")],
  ["panorama sopra e percorso sotto", streetView.includes("grid-template-rows:50% 50%") && streetView.indexOf("hera-sv-panorama-wrap") < streetView.indexOf("hera-sv-route-panel")],
  ["percorso dalla posizione operatore", streetView.includes("const operatorPositionPromise = getOperatorPosition()") && streetView.includes("routeStartKind = operatorCoords ? 'operator' : 'panorama'")],
  ["asset Street View aggiornato", serviceWorker.includes("street-view-cards.js?v=20260830-tree-route1")],
  ["asset CSS aggiornato", html.includes("tree-search.css?v=20260830-tree-route1") && serviceWorker.includes("tree-search.css?v=20260830-tree-route1")],
  ["asset JS aggiornato", html.includes("tree-search.js?v=20260830-tree-route1") && serviceWorker.includes("tree-search.js?v=20260830-tree-route1")],
  ["nessuna operazione Firestore", !/firestore|\.collection\(|\.doc\(/i.test(source)]
];

let failed = false;
for (const [name, passed] of checks) {
  console.log(`${passed ? "OK" : "FAIL"} ${name}`);
  failed ||= !passed;
}

if (failed) process.exit(1);

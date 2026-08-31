#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const read = (file) => fs.readFileSync(path.join(root, file), "utf8");
const source = read("tree-work-orders.js");
const app = read("app.js");
const treeSearch = read("tree-search.js");
const css = read("tree-search.css");
const html = read("index.html");
const serviceWorker = read("sw.js");

const checks = [
  [source.includes('const COMMESSA_ID = "potature-abbattimenti"'), "ID stabile della commessa speciale"],
  [source.includes('const COMMESSA_NAME = "Potature Abbattimenti"'), "nome della commessa speciale"],
  [source.includes('.doc(COMMESSA_ID)') && source.includes('.collection("impianti").doc(plantId)'), "salvataggio nel percorso operativo standard"],
  [source.includes("firstSixDetails(context)") && source.includes(".slice(0, 6)"), "primi sei dettagli del Catasto"],
  [source.includes("schedaParziale: true") && source.includes("campiDaCompletare: true"), "scheda iniziale esplicitamente parziale"],
  [source.includes('stato: "da_fare", done: false, doneAt: null, doneBy: ""') && source.includes("const timestamps = existing"), "stato FATTO preservato negli aggiornamenti"],
  [source.includes("hasPayloadChanges(existing, payload)") && source.includes("if (writes) await batch.commit()"), "nessuna scrittura duplicata se i dati non cambiano"],
  [source.includes("messaggioWhazzupAlbero: dedicatedMessage") && source.includes("🌳 SCHEDA POTATURE / ABBATTIMENTI"), "messaggio dedicato salvato con il cantiere"],
  [source.includes("📍 Naviga verso l’albero:") && source.includes("dettagliCatastoPrimiSei"), "navigazione e campi catasto nel messaggio"],
  [app.includes("function getImpiantoVisibleNote(impianto)") && app.includes("return String(impianto?.noteOperatore || \"\").trim();"), "nota operatore separata dal messaggio tecnico nella scheda"],
  [app.includes("function buildTreeWorkOrderCardDetailsMarkup(impianto)") && app.includes('class="tree-work-order-card-details"'), "primi sei dati dell'albero organizzati in un riquadro dedicato"],
  [app.includes("${treeWorkOrderDetailsMarkup}") && app.includes("${escapeHTML(visibleNoteText)}"), "scheda operativa senza testo tecnico o URL di navigazione"],
  [treeSearch.includes("tree-work-order-open") && treeSearch.includes("window.HeraTreeWorkOrders"), "pulsante collegato alla scheda albero"],
  [css.includes(".tree-work-order-panel") && css.includes(".tree-work-order-prefill"), "interfaccia responsive della scheda"],
  [html.includes("tree-work-orders.js?v=20260831-potature1"), "modulo caricato nell’app"],
  [serviceWorker.includes("tree-work-orders.js?v=20260831-potature1"), "modulo disponibile nella PWA"],
  [!source.includes(".onSnapshot(") && !source.includes("setInterval(") && !source.includes("watchPosition("), "nessun listener o polling aggiunto"],
  [!/(?:function\s+markImpiantoDone|window\.buildImpiantoWhatsAppPayload\s*=|window\.openWhatsApp\s*=|safeOpenWhatsAppMessage\s*=)/.test(source), "nessuna interferenza con FATTO o Whazzup"]
];

const failed = checks.filter(([ok]) => !ok);
checks.forEach(([ok, label]) => console.log(`${ok ? "OK" : "FAIL"} ${label}`));
if (failed.length) process.exit(1);
console.log("✅ Potature Abbattimenti: scheda parziale, cantiere operativo e messaggio dedicato verificati.");

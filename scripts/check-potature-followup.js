#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const workflow = require(path.join(root, "potature-followup.js"));
const app = fs.readFileSync(path.join(root, "app.js"), "utf8");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const css = fs.readFileSync(path.join(root, "style.css"), "utf8");
const serviceWorker = fs.readFileSync(path.join(root, "sw.js"), "utf8");
const source = {
  id: "tree-bologna-103923",
  idSap: "ALB-BOLOGNA-103923",
  denominazione: "ALBERO #103923 — Prunus cerasifera",
  comune: "Bologna",
  indirizzo: "Navile",
  gpsY: 44.5191,
  gpsX: 11.343,
  potatureAbbattimenti: true,
  numeroPunto: "103923",
  dettagliCatastoPrimiSei: [{ etichetta: "Numero punto", valore: "103923" }]
};

assert.equal(workflow.commessaId, "potature-abbattimenti");
assert.equal(workflow.isOriginal(source), true);
assert.equal(workflow.phaseFor({ potatureFase: "raccolta" }), "raccolta");
assert.equal(workflow.isOriginal({ ...source, potatureFase: "ceppi" }), false, "una scheda derivata non deve riaprire il form");

const tasks = workflow.buildTasks(source, { raccolta: "ragno", ceppi: "t15" }, {
  operatorUid: "user-1",
  operatorName: "Mario",
  timestamp: "SERVER_TIMESTAMP"
});
assert.equal(tasks.length, 2);
assert.notEqual(tasks[0].id, tasks[1].id);
assert.equal(tasks[0].id, workflow.taskDocumentId(source, "raccolta"), "ID Raccolta deterministico");
assert.equal(tasks[0].payload.potatureMetodoLabel, "Con ragno");
assert.equal(tasks[0].payload.lavorazioniRichieste, "Mucchia con ragno");
assert.equal(tasks[1].payload.potatureMetodoLabel, "Con T15");
assert.equal(tasks[1].payload.lavorazioniRichieste, "Ceppo con t15");
assert.equal(tasks[0].payload.denominazione, source.denominazione, "stesso impianto copiato nella vista");
assert.equal(tasks[0].payload.gpsY, source.gpsY);
assert.equal(tasks[0].payload.potatureOrigineId, source.id);
assert.equal(Object.hasOwn(tasks[0].payload, "done"), false, "il salvataggio esterno non modifica lo stato FATTO");
assert.equal(Object.hasOwn(tasks[0].payload, "doneAt"), false, "il salvataggio esterno non modifica data FATTO");
assert.equal(Object.hasOwn(tasks[0].payload, "doneBy"), false, "il salvataggio esterno non modifica operatore FATTO");
assert.deepEqual(workflow.buildTasks(source, { raccolta: "", ceppi: "" }), [], "entrambe le scelte sono facoltative");

const previous = workflow.existingSelections(source, tasks.map((task) => task.payload));
assert.deepEqual(previous, { raccolta: "ragno", ceppi: "t15" }, "le scelte salvate vengono riproposte nel form");

assert.match(app, /impiantiViewMode === "raccolta" \|\| impiantiViewMode === "ceppi"/);
assert.match(app, /PREPARA FINE/);
assert.match(app, /mappedImpianti = currentImpianti\.filter/);
assert.match(html, /id="view-raccolta-btn"/);
assert.match(html, /id="view-ceppi-btn"/);
assert.match(html, /potature-followup\.js\?v=20260902-special-immediate-move1/);
assert.match(serviceWorker, /potature-followup\.js\?v=20260902-special-immediate-move1/);
assert.match(css, /\.potature-followup-modal/);

const runtimeSource = fs.readFileSync(path.join(root, "potature-followup.js"), "utf8");
assert.match(runtimeSource, /window\.HeraSpecialTerminato/,
  "il flusso TERMINATO speciale deve essere installato separatamente");
assert.match(runtimeSource, /specialTerminatoAt/,
  "TERMINATO registra una data separata");
assert.match(runtimeSource, /specialTerminatoBy/,
  "TERMINATO registra un operatore separato");
assert.match(runtimeSource, /finishedButton\.textContent = "✅ Finiti"/,
  "Fatti diventa Finiti nelle commesse speciali");
assert.match(runtimeSource, /programButton\.textContent = "🛠️ In programma"/,
  "Da fare diventa In programma nelle commesse speciali");
assert.match(runtimeSource, /exportButton\.textContent = "📤 Esporta finiti"/,
  "Esporta fatti diventa Esporta finiti nelle commesse speciali");
assert.match(runtimeSource, /currentPlants\(\)\.filter\(isTerminated\)/,
  "l'esportazione usa gli impianti Finiti già caricati");
assert.match(runtimeSource, /XLSX\.writeFile\(workbook, `riepilogo_impianti_/,
  "l'esportazione mantiene il download Excel del vecchio sistema");
assert.match(runtimeSource, /navigateButton\.dataset\.actionKey = "navigate"/,
  "NAVIGA mantiene posizione e stile del flusso standard nella vista Finiti");
assert.match(runtimeSource, /statusButton\.className = "btn special-finished-status-btn"/,
  "lo stato TERMINATO occupa la posizione del pulsante operativo nella vista Finiti");
assert.match(runtimeSource, /typeof matchesImpiantoSearch === "function"/,
  "la ricerca continua a filtrare anche la vista Finiti");
assert.match(runtimeSource, /typeof distanceFromUser === "function"/,
  "la vista Finiti mantiene l'ordinamento per distanza");
assert.match(runtimeSource, /showFinishedList\(\)/,
  "TERMINATO sposta immediatamente il cantiere nella vista Finiti");
assert.match(runtimeSource, /heraSpecialTerminatoPendingV2/,
  "TERMINATO mantiene una coda offline separata e persistente");
assert.match(runtimeSource, /persistActionWithRetry/,
  "TERMINATO riprova il salvataggio senza coinvolgere FATTO");
assert.match(runtimeSource, /specialTerminatoByEmail/,
  "TERMINATO conserva anche l’email dell’operatore");
assert.match(runtimeSource, /specialDataEsecuzione/,
  "TERMINATO conserva la data di esecuzione separata");
assert.match(runtimeSource, /specialOraEsecuzione/,
  "TERMINATO conserva l’ora di esecuzione separata");
assert.match(runtimeSource, /window\.addEventListener\("online"/,
  "la coda TERMINATO riparte automaticamente al ritorno della rete");
assert.match(runtimeSource, /publishGlobalNotificationEvent\("impianto-done"/,
  "TERMINATO pubblica la notifica operativa già prevista dall’app");
assert.match(runtimeSource, /logActivity\("pressione_terminato"/,
  "TERMINATO viene registrato nello storico attività");
assert.match(app, /data-map-popup-action='special-terminate'/,
  "le mappe collegano le commesse speciali al nuovo pulsante TERMINATO");
assert.match(app, /window\.HeraSpecialTerminato\?\.terminateFromMap/,
  "il pulsante mappa usa esclusivamente il sistema speciale");
assert.doesNotMatch(runtimeSource, /view-special-terminated-btn/,
  "il pulsante Terminati separato è stato eliminato");
assert.match(runtimeSource, /special-core-action-hidden/,
  "le azioni storiche vengono soltanto nascoste nelle commesse speciali");
assert.doesNotMatch(runtimeSource, /\[TERMINATED_FIELD\]:\s*true[\s\S]{0,300}\bdone\s*:/,
  "TERMINATO non deve scrivere lo stato FATTO");
assert.match(runtimeSource, /reference\.doc\(id\)\.get\(\)/,
  "una sola verifica mirata controlla tutti i documenti appena salvati");
assert.doesNotMatch(runtimeSource, /onSnapshot|setInterval|watchPosition/, "il nuovo flusso non aggiunge listener o polling");
assert.doesNotMatch(runtimeSource, /markImpiantoDone|handleImpiantoWhatsAppClick|openWhatsApp/, "il nuovo flusso non richiama né aggira FATTO/WHAZZUP");

console.log("✅ Potature Abbattimenti e COBO: In programma/Finiti, TERMINATO separato e isolamento FATTO verificati.");

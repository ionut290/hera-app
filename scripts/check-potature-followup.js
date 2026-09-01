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
assert.equal(Object.hasOwn(tasks[0].payload, "done"), false, "il salvataggio delle attività derivate non modifica lo stato completato");
assert.equal(Object.hasOwn(tasks[0].payload, "doneAt"), false, "il salvataggio delle attività derivate non modifica la data di completamento");
assert.equal(Object.hasOwn(tasks[0].payload, "doneBy"), false, "il salvataggio delle attività derivate non modifica l’operatore di completamento");
assert.deepEqual(workflow.buildTasks(source, { raccolta: "", ceppi: "" }), [], "entrambe le scelte sono facoltative");

const previous = workflow.existingSelections(source, tasks.map((task) => task.payload));
assert.deepEqual(previous, { raccolta: "ragno", ceppi: "t15" }, "le scelte salvate vengono riproposte nel form");

assert.match(app, /impiantiViewMode === "raccolta" \|\| impiantiViewMode === "ceppi"/);
assert.match(app, /PREPARA FINE/);
assert.match(app, /mappedImpianti = currentImpianti\.filter/);
assert.match(html, /id="view-raccolta-btn"/);
assert.match(html, /id="view-ceppi-btn"/);
assert.match(html, /potature-followup\.js\?v=20260831-potature-followup1/);
assert.match(serviceWorker, /potature-followup\.js\?v=20260831-potature-followup1/);
assert.match(css, /\.potature-followup-modal/);

const runtimeSource = fs.readFileSync(path.join(root, "potature-followup.js"), "utf8");
assert.doesNotMatch(runtimeSource, /await\s+[^;\n]*\.get\s*\(/, "TERMINATO non aggiunge letture Firestore");
assert.doesNotMatch(runtimeSource, /onSnapshot|setInterval|watchPosition/, "TERMINATO non aggiunge listener Firestore, polling o tracking posizione");
assert.doesNotMatch(runtimeSource, /markImpiantoDone|handleImpiantoWhatsAppClick|forceMoveImpiantoToFatti|openWhatsApp/, "TERMINATO non richiama né aggira le funzioni protette FATTO/WHAZZUP");
assert.match(runtimeSource, /completionMode:\s*"TERMINATO_SPECIAL"/, "il completamento speciale deve essere tracciabile");
assert.match(runtimeSource, /value\.includes\("sfalcio"\)\s*&&\s*value\.includes\("cobo"\)/, "Sfalcio COBO deve essere riconosciuto anche con ID dinamico");
assert.match(runtimeSource, /TERMINATO_ACTION\s*=\s*"special-terminato"/, "deve esistere una azione TERMINATO separata");
assert.match(runtimeSource, /handleCompletedImpiantoWhatsAppClick/, "dopo il salvataggio deve essere riutilizzato solo il Whazzup dell’impianto già completato");
assert.match(runtimeSource, /style\.setProperty\("display",\s*"none",\s*"important"\)/, "il vecchio FATTO deve essere completamente invisibile");
assert.match(runtimeSource, /insertBefore\(button,\s*legacyTarget\)/, "TERMINATO deve prendere esattamente la posizione del vecchio FATTO");
assert.match(runtimeSource, /dataset\.replacesAction\s*=\s*"fatto"/, "TERMINATO deve dichiarare esplicitamente che sostituisce FATTO");
assert.doesNotMatch(runtimeSource, /Ora puoi premere FATTO|premi il pulsante verde <strong>FATTO<\/strong>/, "il form Potature non deve più invitare a usare FATTO");
assert.match(runtimeSource, /\.action-icon-btn\[data-action-key="whatsapp"\].*label\.includes\("fatto"\)/s, "un eventuale WhatsApp separato non deve essere nascosto se non contiene FATTO");

console.log("✅ Potature Abbattimenti + Sfalcio COBO: FATTO invisibile, TERMINATO nella stessa posizione e isolamento FATTO verificati.");

#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const assert = require("node:assert/strict");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "..", "functions", "shared-calendar-view.js"), "utf8");
const main = fs.readFileSync(path.resolve(__dirname, "..", "functions", "main.js"), "utf8");

assert.match(source, /oreReports\/\{documentId\}/, "Manca il trigger oreReports");
assert.match(source, /oreApprovalRequests\/\{documentId\}/, "Manca il trigger oreApprovalRequests");
assert.match(source, /commesse\/\{commessaId\}\/impianti\/\{itemId\}/, "Manca il trigger impianti");
assert.match(source, /commesse\/\{commessaId\}\/lavorazioni\/\{itemId\}/, "Manca il trigger lavorazioni");
assert.match(source, /runTransaction/, "L'aggiornamento condiviso deve essere transazionale");
assert.match(source, /calendario__\$\{month\}/, "Il documento deve essere mensile");
assert.match(source, /payload-too-large/, "Manca il limite di sicurezza del payload");
assert.match(source, /beforeMonth/, "Deve gestire lo spostamento tra mesi");
assert.match(source, /afterMonth/, "Deve gestire creazione, modifica ed eliminazione");
assert.match(source, /activities: Array\.isArray\(existingPayload\?\.activities\)/, "Le ore devono conservare le attività già aggregate");
assert.match(source, /reports: Array\.isArray\(existingPayload\?\.reports\)/, "Le attività devono conservare le ore già aggregate");
assert.match(source, /compactActivity/, "Le attività devono essere ridotte ai soli campi amministrativi");
assert.match(source, /getAdministrativeCalendarMonth/, "Manca il recupero mensile controllato");
assert.match(source, /collectionGroup\(sourceCollection\)/, "Il recupero deve usare query mensili collectionGroup");
assert.match(source, /where\("dataEsecuzione", ">=", fromDate\)/, "Il recupero deve essere limitato all'inizio del mese");
assert.match(source, /where\("dataEsecuzione", "<", toDate\)/, "Il recupero deve essere limitato alla fine del mese");
assert.match(source, /collection\("oreReports"\)[\s\S]*?where\("date", ">=", fromDate\)[\s\S]*?where\("date", "<", toDate\)/, "Il recupero mensile deve includere le ore già inserite");
assert.match(source, /collection\("oreApprovalRequests"\)[\s\S]*?where\("date", ">=", fromDate\)[\s\S]*?where\("date", "<", toDate\)/, "Il recupero mensile deve includere le ore in approvazione");
assert.match(source, /reports,\s*activities,\s*recoveredReports:/, "Il recupero deve restituire ore e attività al calendario");
assert.match(source, /if \(!request\.auth\)/, "Il recupero deve richiedere autenticazione");
assert.match(source, /MAX_RECOVERED_ACTIVITIES/, "Manca il limite del recupero mensile");
assert.match(main, /shared-calendar-view/, "Le funzioni non sono esportate da main.js");

const moduleStub = { exports: {} };
vm.runInNewContext(source, {
  require: (id) => {
    if (id === "crypto") return require("node:crypto");
    if (id === "firebase-admin") return {};
    if (id === "firebase-functions/v2/firestore") return { onDocumentWritten: (_options, handler) => handler };
    if (id === "firebase-functions/v2/https") return {
      onCall: (_options, handler) => handler,
      HttpsError: class HttpsError extends Error {}
    };
    throw new Error(`Modulo inatteso: ${id}`);
  },
  module: moduleStub,
  exports: moduleStub.exports,
  Buffer,
  Date,
  Set,
  Number,
  String,
  Object,
  Intl,
  console
}, { filename: "shared-calendar-view.js" });

const helpers = moduleStub.exports.__test;
const activity = helpers.compactActivity("lavorazioni", "c1", "w1", {
  stato: "FATTO",
  dataEsecuzione: "2026-08-26",
  oraEsecuzione: "08:30",
  impiantoId: "i1",
  tipologiaLavorazione: "Sfalcio",
  operatoreNome: "Mario Rossi",
  totale: "1.234,50"
});
assert.equal(activity.date, "2026-08-26");
assert.equal(activity.amount, 1234.5);
assert.equal(activity.kind, "lavorazione");
assert.equal(helpers.compactActivity("lavorazioni", "c1", "w-price", {
  stato: "FATTO", dataEsecuzione: "2026-08-26", codiceVocePrezzo: "A1",
  quantita: "2,5", unitaMisura: "M2", prezzoRibassato: "79,20"
}).amount, 198);
assert.equal(helpers.compactActivity("lavorazioni", "c1", "w-ac", {
  stato: "FATTO", dataEsecuzione: "2026-08-26", codiceVocePrezzo: "A2",
  unitaMisura: "AC", prezzoBase: 100, percentualeRibasso: 0.01
}).amount, 99);
assert.equal(helpers.activityDateKeyFromData({ doneAt: { toDate: () => new Date("2026-08-25T22:30:00Z") } }), "2026-08-26");
assert.equal(helpers.nextMonthKey("2026-12"), "2027-01");
assert.equal(helpers.compactActivity("lavorazioni", "c1", "w2", { stato: "DA FARE" }), null);

const hoursUpdate = helpers.buildNextPayload({ activities: [activity] }, "2026-08", "oreReports", "r1", { date: "2026-08-26" });
assert.equal(hoursUpdate.activities.length, 1);
const activityUpdate = helpers.buildNextActivityPayload({ reports: hoursUpdate.reports }, "2026-08", "lavorazioni", "c1", "w1", {
  stato: "FATTO", dataEsecuzione: "2026-08-26", impiantoId: "i1", totale: 10
});
assert.equal(activityUpdate.reports.length, 1);
assert.equal(activityUpdate.activities.length, 1);

console.log("✅ Trigger automatico collegato a ore, impianti e lavorazioni FATTO.");
console.log("✅ Recupero storico di ore e attività limitato al mese richiesto, autenticato e salvato nel riepilogo condiviso.");

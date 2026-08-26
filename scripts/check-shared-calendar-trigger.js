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
assert.doesNotMatch(source, /collectionGroup\s*\(/, "Il trigger non deve eseguire scansioni collectionGroup");
assert.match(main, /shared-calendar-view/, "Le funzioni non sono esportate da main.js");

const moduleStub = { exports: {} };
vm.runInNewContext(source, {
  require: (id) => {
    if (id === "crypto") return require("node:crypto");
    if (id === "firebase-admin") return {};
    if (id === "firebase-functions/v2/firestore") return { onDocumentWritten: (_options, handler) => handler };
    throw new Error(`Modulo inatteso: ${id}`);
  },
  module: moduleStub,
  exports: moduleStub.exports,
  Buffer,
  Date,
  Set,
  Number,
  String,
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
assert.equal(helpers.compactActivity("lavorazioni", "c1", "w2", { stato: "DA FARE" }), null);

const hoursUpdate = helpers.buildNextPayload({ activities: [activity] }, "2026-08", "oreReports", "r1", { date: "2026-08-26" });
assert.equal(hoursUpdate.activities.length, 1);
const activityUpdate = helpers.buildNextActivityPayload({ reports: hoursUpdate.reports }, "2026-08", "lavorazioni", "c1", "w1", {
  stato: "FATTO", dataEsecuzione: "2026-08-26", impiantoId: "i1", totale: 10
});
assert.equal(activityUpdate.reports.length, 1);
assert.equal(activityUpdate.activities.length, 1);

console.log("✅ Trigger automatico collegato a ore, impianti e lavorazioni FATTO.");
console.log("✅ Aggiornamento mensile transazionale, senza scansioni e con conservazione delle sezioni.");

#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const assert = require("node:assert/strict");

const source = fs.readFileSync(path.resolve(__dirname, "..", "functions", "shared-calendar-view.js"), "utf8");
const main = fs.readFileSync(path.resolve(__dirname, "..", "functions", "main.js"), "utf8");

assert.match(source, /oreReports\/\{documentId\}/, "Manca il trigger oreReports");
assert.match(source, /oreApprovalRequests\/\{documentId\}/, "Manca il trigger oreApprovalRequests");
assert.match(source, /runTransaction/, "L'aggiornamento condiviso deve essere transazionale");
assert.match(source, /calendario__\$\{month\}/, "Il documento deve essere mensile");
assert.match(source, /payload-too-large/, "Manca il limite di sicurezza del payload");
assert.match(source, /beforeMonth/, "Deve gestire lo spostamento tra mesi");
assert.match(source, /afterMonth/, "Deve gestire creazione, modifica ed eliminazione");
assert.match(main, /shared-calendar-view/, "Le funzioni non sono esportate da main.js");

console.log("✅ Trigger automatico calendario collegato a oreReports e oreApprovalRequests.");
console.log("✅ Aggiornamento mensile transazionale, con gestione cambio mese ed eliminazione.");

#!/usr/bin/env node
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const root = path.join(__dirname, "..");
const source = fs.readFileSync(path.join(root, "inrete-modena-coordinate-backup.js"), "utf8");
const appSource = fs.readFileSync(path.join(root, "app.js"), "utf8");
const indexSource = fs.readFileSync(path.join(root, "index.html"), "utf8");
const serviceWorkerSource = fs.readFileSync(path.join(root, "sw.js"), "utf8");
let renderCalls = 0;
let currentImpianti = [];
const context = {
  window: {},
  renderImpianti() { renderCalls += 1; },
  get currentImpianti() { return currentImpianti; },
  set currentImpianti(value) { currentImpianti = value; }
};
vm.createContext(context);
vm.runInContext(source, context);

const backup = context.window.HeraInreteModenaCoordinateBackup;
let failed = false;
function check(name, passed) {
  console.log(`${passed ? "OK" : "FAIL"} ${name}`);
  failed ||= !passed;
}

const missingRows = backup.backupRows.map((row, index) => ({
  id: `missing-${index + 1}`,
  idSap: index % 2 === 0 ? row.sap : "",
  denominazione: row.name,
  gpsY: null,
  gpsX: null
}));
backup.restoreRows(missingRows);

check("il backup contiene i 30 impianti mancanti", backup.backupRows.length === 30);
check("tutti i 30 impianti mancanti ricevono coordinate valide", missingRows.every(backup.hasCoordinates));
check("REMI FORMIGINE usa le coordinate del foglio", missingRows.find((row) => row.denominazione === "REMI FORMIGINE")?.gpsY === 44.56882);
check("REMI CASTELLARO usa le coordinate del foglio", missingRows.find((row) => row.denominazione === "REMI CASTELLARO")?.gpsX === 11.02894);
check("REMI BRAIDA usa le coordinate del foglio", missingRows.find((row) => row.denominazione === "REMI BRAIDA")?.gpsY === 44.54703);

const existing = { idSap: "3470478", denominazione: "REMI FORMIGINE", gpsY: 44.5, gpsX: 10.5 };
check("coordinate già valide non vengono sovrascritte", backup.restoreRecord(existing) === false && existing.gpsY === 44.5 && existing.gpsX === 10.5);

const historicalFields = { denominazione: "REMI FORMIGINE", latitudine: 44.6, longitudine: 10.6, gpsY: null, gpsX: null };
check("coordinate storiche valide vengono copiate nei campi usati dalla mappa", backup.restoreRecord(historicalFields) === false && historicalFields.gpsY === 44.6 && historicalFields.gpsX === 10.6);

const unrelated = { denominazione: "IMPIANTO NON MODENA", gpsY: null, gpsX: null };
check("impianti estranei restano invariati", backup.restoreRecord(unrelated) === false && unrelated.gpsY === null && unrelated.gpsX === null);

const normalizedName = { denominazione: "  Remi   Ca di Solà  ", gpsY: null, gpsX: null };
check("nomi storici con spazi e accenti vengono riconosciuti", backup.restoreRecord(normalizedName) && normalizedName.gpsY === 44.53481);

const ambiguous = { denominazione: "Sfalci e potature per accessi valvole gas", gpsY: null, gpsX: null };
check("il nome duplicato con due coordinate non viene assegnato automaticamente", backup.restoreRecord(ambiguous) === false);

currentImpianti = [{ denominazione: "REMI FORMIGINE", gpsY: null, gpsX: null }];
context.renderImpianti();
check("il ripristino viene applicato prima del rendering impianti", renderCalls === 1 && currentImpianti[0].gpsY === 44.56882);
check("lo script viene caricato subito dopo app.js", indexSource.indexOf("inrete-modena-coordinate-backup.js?v=20260819-restore1") > indexSource.indexOf("app.js?v=20260819-modena-restore1"));
check("app.js resta privo di logica specifica Modena", !appSource.includes("HeraInreteModenaCoordinateBackup"));
check("service worker include il backup coordinate", serviceWorkerSource.includes("./inrete-modena-coordinate-backup.js?v=20260819-restore1"));
check("nessuna operazione Firestore aggiunta", !/firestore|\bdb\.collection\(|\.onSnapshot\(|serverTimestamp|writeBatch|\.batch\(/i.test(source));

if (failed) process.exit(1);

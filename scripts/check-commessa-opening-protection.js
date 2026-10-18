"use strict";

const { execFileSync, spawnSync } = require("node:child_process");
const fs = require("node:fs");

const LOCKED_FILE = "commessa-navigation-repair.js";
const EXPECTED_GIT_BLOB_SHA = "56f89add5910b30f71c6e37a80539ed767e21aa6";

function fail(message) {
  console.error(`\n[COMMESSA OPENING GUARD] ${message}`);
  process.exit(1);
}

if (!fs.existsSync(LOCKED_FILE)) {
  fail(`${LOCKED_FILE} non esiste. Il flusso protetto di apertura commessa non può essere rimosso.`);
}

const currentSha = execFileSync("git", ["hash-object", LOCKED_FILE], { encoding: "utf8" }).trim();
if (currentSha !== EXPECTED_GIT_BLOB_SHA) {
  fail(
    `${LOCKED_FILE} è stato modificato.\n` +
    `Versione protetta: ${EXPECTED_GIT_BLOB_SHA}\n` +
    `Versione trovata:  ${currentSha}\n\n` +
    "Questo file gestisce apertura commessa, caricamento/caching impianti e lifecycle dei listener. " +
    "Le modifiche ordinarie devono essere fatte altrove. Per cambiare intenzionalmente questo flusso occorre prima validare la nuova versione e aggiornare esplicitamente questa baseline."
  );
}

const source = fs.readFileSync(LOCKED_FILE, "utf8");
const requiredMarkers = [
  "originalSelectCommessa",
  "forceCommessaNavigation",
  "HeraImpiantiPersistentCache",
  "subscribeImpiantiWithPersistentCache",
  "renderImpiantiAfterRemoteSyncWithPersistentCache",
  "HeraLazyCommessaStats",
  "HeraImpiantiListenerLifecycleGuard"
];
for (const marker of requiredMarkers) {
  if (!source.includes(marker)) fail(`Manca il marker critico ${marker}.`);
}

const checks = [
  "scripts/check-critical-flows.js",
  "scripts/check-commessa-impianti-menu.js",
  "scripts/check-commessa-stats-cache.js",
  "scripts/check-commesse-impianti-integrity.js",
  "scripts/check-impianti-listener-lifecycle.js",
  "scripts/check-startup-essential-data-fallback.js"
];

for (const check of checks) {
  if (!fs.existsSync(check)) fail(`Test di regressione mancante: ${check}`);
  const result = spawnSync(process.execPath, [check], { stdio: "inherit" });
  if (result.status !== 0) fail(`Test fallito: ${check}`);
}

console.log("\n[COMMESSA OPENING GUARD] OK: apertura commessa e caricamento dati protetti.");

#!/usr/bin/env node
const fs = require("fs");
const path = require("path");

const runtimePath = path.join(__dirname, "..", "squadra-current-save-sync.js");
const source = fs.readFileSync(runtimePath, "utf8");
const guardMarker = "if (window.HeraFattoPersistenceReadGuard?.installed) return;";
const guardStart = source.indexOf(guardMarker);

function fail(message) {
  console.error(`❌ ${message}`);
  process.exitCode = 1;
}

function pass(message) {
  console.log(`✅ ${message}`);
}

function requireIncludes(scope, needle, message) {
  if (!scope.includes(needle)) fail(message);
  else pass(message);
}

if (guardStart < 0) {
  fail("guard verifica FATTO presente");
  process.exit(process.exitCode || 1);
}

pass("guard verifica FATTO presente");
const guard = source.slice(guardStart);
requireIncludes(guard, "isImpiantoPersistedAsDone = async function isImpiantoPersistedAsDoneFromCache", "verifica FATTO sostituita dopo app.js");
requireIncludes(guard, '.get({ source: "cache" })', "verifica FATTO usa esclusivamente la cache Firestore");
requireIncludes(guard, "getLocalImpianti(commessaId)", "fallback usa i dati già consegnati dal listener impianti");
requireIncludes(guard, 'mode: "cache-listener-no-network-get"', "modalità diagnostica dichiara assenza di get di rete");

if (/\.get\(\s*\)/.test(guard)) fail("il guard FATTO non deve contenere doc.get() di rete");
else pass("il guard FATTO non contiene doc.get() di rete");

if (/source:\s*["']server["']/.test(guard)) fail("il guard FATTO non deve forzare letture server");
else pass("il guard FATTO non forza letture server");

if (process.exitCode) process.exit(process.exitCode);
console.log("✅ Controlli guard lettura FATTO completati.");

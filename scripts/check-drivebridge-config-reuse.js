#!/usr/bin/env node
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const app = fs.readFileSync("app.js", "utf8");
const registry = fs.readFileSync("registry-google-sheet-sync.js", "utf8");
assert.match(app, /function subscribeDriveBridge\(\)[\s\S]*?HeraDriveBridgeConfigShared\?\.publish[\s\S]*?\.publish\(data \|\| \{\}\)/, "Il listener Drive deve pubblicare il proprio snapshot nel ponte condiviso");
assert.match(registry, /function getSharedDriveBridgeConfigState\(\)/, "Deve esistere il ponte condiviso della configurazione Drive");
assert.match(registry, /async function waitForSharedDriveBridgeConfig\(timeoutMs = 900\)/, "Il riuso deve avere un timeout breve e finito");
assert.match(registry, /if \(!force && sharedDriveBridgeConfig\.hasSnapshot\)/, "La cache del listener deve avere precedenza sui get non forzati");
assert.match(registry, /waitForSharedDriveBridgeConfig\(\)[\s\S]*?sharedData !== null[\s\S]*?configDocRef\(\)\.get\(\)/, "Se il listener non consegna dati deve restare il get Firestore di fallback");
assert.match(registry, /if \(!force\)[\s\S]*?return configDocumentPromise;[\s\S]*?configDocumentPromise = configDocRef\(\)\.get\(\)/, "I caricamenti force devono continuare a leggere Firestore direttamente");
assert.doesNotMatch(app, /HeraDriveBridgeConfigShared[\s\S]{0,300}\.onSnapshot\(/, "Il ponte non deve creare un nuovo listener Firestore");
console.log("✅ appConfig/driveBridge riusa il listener esistente prima del get di fallback.");
console.log("✅ Nessun nuovo listener viene creato e force=true conserva la lettura server.");

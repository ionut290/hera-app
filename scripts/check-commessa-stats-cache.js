#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const code = fs.readFileSync("commessa-stats-cache-optimizer.js", "utf8");
const index = fs.readFileSync("index.html", "utf8");

const appScriptIndex = index.indexOf("app.js?v=");
const statsScriptIndex = index.indexOf("commessa-stats-cache-optimizer.js");
const sharedClientIndex = index.indexOf("shared-static-views-client.js");
const navigationScriptIndex = index.indexOf("commessa-navigation-repair.js");
assert.ok(appScriptIndex >= 0, "app.js script tag missing");
assert.ok(statsScriptIndex >= 0, "Stats optimizer script tag missing");
assert.ok(sharedClientIndex >= 0, "Shared static views client script tag missing");
assert.ok(navigationScriptIndex >= 0, "Navigation repair script tag missing");
assert.ok(appScriptIndex < statsScriptIndex, "Optimizer must load after app.js definitions");
assert.ok(statsScriptIndex < sharedClientIndex, "Optimizer must load before Light Startup captures subscribeStatsForCommesse");
assert.ok(statsScriptIndex < navigationScriptIndex, "Optimizer must load before navigation repair");
assert.match(code, /heraCommessaStatsCacheV1:/);
assert.match(code, /impiantoChangeIndex/);
assert.match(code, /where\("changedAt", ">", markerDate\)/);
assert.match(code, /orderBy\("changedAt", "asc"\)/);
assert.match(code, /\.doc\(impiantoId\)\s*;\s*const snap = await ref\.get\(\)/s);
assert.match(code, /fallbackFullLoad/);
assert.match(code, /collection\("impianti"\)\.get\(\)/);
assert.match(code, /readImpiantiCache/);
assert.match(code, /calculateImpiantiStats/);
assert.match(code, /recalculateCommessaWorkSummaries/);

// Sono ammessi Map.set/localStorage.setItem. Vietiamo invece mutazioni Firestore
// nel nuovo ottimizzatore: deve essere strettamente read-only lato server.
assert.doesNotMatch(code, /\.collection\([^\n]+\)\.add\(/, "Optimizer must not add Firestore documents");
assert.doesNotMatch(code, /\.doc\([^\n]+\)\.(?:set|update|delete)\(/, "Optimizer must not mutate Firestore documents");
assert.doesNotMatch(code, /writeBatch|runTransaction|FieldValue\./, "Optimizer must not use Firestore write APIs");
assert.doesNotMatch(code, /FATTO|fattoVisualEvidence|WhatsApp|WHAZZUP|whazzup/i, "Protected FATTO/WhatsApp flows must stay outside this optimizer");

console.log("Commessa stats cache safety checks passed");

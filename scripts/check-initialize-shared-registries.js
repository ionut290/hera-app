#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const assert = require("node:assert/strict");

const workflow = fs.readFileSync(".github/workflows/initialize-shared-registries.yml", "utf8");
const core = fs.readFileSync("functions/shared-operational-views.js", "utf8");
const runner = fs.readFileSync("functions/scripts/initialize-shared-registries.js", "utf8");

assert.match(workflow, /workflow_dispatch:/);
assert.match(workflow, /FIREBASE_SERVICE_ACCOUNT_HERA_APP_6CD2B/);
assert.match(workflow, /initialize-shared-registries\.js/);
assert.match(core, /__server = \{ rebuildRegistryView \}/);
assert.match(runner, /INITIALIZE_SHARED_REGISTRIES_RESULT=/);
assert.match(runner, /rebuildRegistryView/);
assert.doesNotMatch(runner, /rebuildSquadreDate|rebuildCalendarMonth/);
assert.doesNotMatch(runner, /httpsCallable|cloudfunctions\.net|fetch\(/);
assert.doesNotMatch(runner, /firestore\(\).*\.(delete|add|update)\(/s);

console.log("✅ Inizializzazione server-side dei soli registri verificata.");

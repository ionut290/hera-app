#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const assert = require("node:assert/strict");

const workflow = fs.readFileSync(".github/workflows/backfill-shared-static-views.yml", "utf8");
const core = fs.readFileSync("functions/shared-operational-views.js", "utf8");
const runner = fs.readFileSync("functions/scripts/backfill-shared-static-views.js", "utf8");

assert.match(workflow, /workflow_dispatch:/);
assert.match(workflow, /default: "2026-08-04"/);
assert.match(workflow, /default: "2026-08"/);
assert.match(workflow, /FIREBASE_SERVICE_ACCOUNT_HERA_APP_6CD2B/);
assert.match(workflow, /backfill-shared-static-views\.js/);
assert.match(core, /__server = \{ rebuildRegistryView, rebuildSquadreDate, rebuildCalendarMonth \}/);
assert.match(runner, /BACKFILL_SHARED_STATIC_VIEWS_RESULT=/);
assert.match(runner, /documentiGiaScritti/);
assert.doesNotMatch(runner, /httpsCallable|cloudfunctions\.net|fetch\(/);
console.log("✅ Workflow server-side senza HTTP/CORS verificato.");

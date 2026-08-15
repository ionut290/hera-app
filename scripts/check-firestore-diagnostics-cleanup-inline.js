#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("firebase-config.js", "utf8");

assert.match(source, /function recoverAbandonedFirestoreDiagnosticListeners\(/);
assert.match(source, /previous-page-listeners-recovered/);
assert.match(source, /page-session-ended-without-unsubscribe/);
assert.match(source, /recoverAbandonedFirestoreDiagnosticListeners\(\);/);
assert.doesNotMatch(source, /firestore-diagnostics-v4-session-cleanup\.js/);
assert.equal(fs.existsSync("firestore-diagnostics-v4-session-cleanup.js"), false, "Il vecchio file cleanup non deve esistere");

console.log("✅ Cleanup diagnostico V4 incorporato nel bootstrap Firebase.");

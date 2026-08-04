#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const assert = require("node:assert/strict");

const source = fs.readFileSync("activity-logs-read-disable.js", "utf8");

assert.match(source, /loadActiveUsersLogs = async function loadActiveUsersLogsDisabled/);
assert.match(source, /activeUsersLogs = \[\]/);
assert.match(source, /Registro attività disattivato/);
assert.match(source, /mode: "reads-disabled"/);
assert.doesNotMatch(source, /db\.|\.collection\(|\.get\(|\.onSnapshot\(|runFirestoreGetWithRetry/);
assert.doesNotMatch(source, /\.delete\(/);

console.log("✅ Letture activityLogs disattivate senza cancellazioni.");

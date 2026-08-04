#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const assert = require("node:assert/strict");

const source = fs.readFileSync("activity-logs-read-disable.js", "utf8");

assert.match(source, /appConfig\/activeCommesse/);
assert.match(source, /non-destructive-listener-filter-v2/);
assert.match(source, /\^commesse\\\/\(\[\^\/\]\+\)\\\/impianti\$/);
assert.match(source, /Le ore storiche restano nel calendario personale/);
assert.match(source, /Listener impianti evitato per commessa disattivata/);
assert.match(source, /firebase\.firestore\.FieldValue\.serverTimestamp/);
assert.match(source, /Firestore non ha confermato lo stato della commessa/);
assert.match(source, /event\.stopImmediatePropagation\(\)/);
assert.match(source, /get\(\{ source: "server" \}\)/);
assert.doesNotMatch(source, /window\.location\.reload/);
assert.doesNotMatch(source, /setTimeout\([^\n]*reload/);

assert.doesNotMatch(source, /\.delete\s*\(/);
assert.doesNotMatch(source, /batch\.delete/);
assert.doesNotMatch(source, /oreReports[^\n]*(delete|where|onSnapshot)/i);
assert.doesNotMatch(source, /oreApprovalRequests[^\n]*(delete|where|onSnapshot)/i);
assert.doesNotMatch(source, /squadreStorico[^\n]*(delete|where|onSnapshot)/i);
assert.doesNotMatch(source, /platformUsers[^\n]*(delete|where|onSnapshot)/i);

console.log("✅ Disattivazione commesse confermata da Firestore senza riavvio automatico.");

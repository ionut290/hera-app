#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const assert = require("node:assert/strict");

const source = fs.readFileSync("activity-logs-read-disable.js", "utf8");

assert.match(source, /appConfig\/activeCommesse/);
assert.match(source, /active-commesse-startup-query-filter-v3/);
assert.match(source, /\^commesse\\\/\(\[\^\/\]\+\)\\\/impianti\$/);
assert.match(source, /isRootCommesseQuery/);
assert.match(source, /firebase\.firestore\.FieldPath\.documentId\(\)/);
assert.match(source, /"in",\s*ids/);
assert.match(source, /loadAllCommesseForManager/);
assert.match(source, /db\.collection\("commesse"\)\.get\(\)/);
assert.match(source, /Le ore storiche restano nel calendario personale/);
assert.match(source, /Listener impianti evitato per commessa disattivata/);
assert.match(source, /Listener commesse iniziale limitato agli ID attivi/);
assert.match(source, /firebase\.firestore\.FieldValue\.serverTimestamp/);
assert.match(source, /window\.location\.reload/);

assert.doesNotMatch(source, /(?:batch|ref|doc|snapshot)\.delete\s*\(/);
assert.doesNotMatch(source, /batch\.delete/);
assert.doesNotMatch(source, /oreReports[^\n]*(delete|where|onSnapshot)/i);
assert.doesNotMatch(source, /oreApprovalRequests[^\n]*(delete|where|onSnapshot)/i);
assert.doesNotMatch(source, /squadreStorico[^\n]*(delete|where|onSnapshot)/i);
assert.doesNotMatch(source, /platformUsers[^\n]*(delete|where|onSnapshot)/i);

console.log("✅ Avvio limitato alle commesse attive; storico ore e dati restano non distruttivi.");

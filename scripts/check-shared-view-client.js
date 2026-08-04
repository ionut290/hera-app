#!/usr/bin/env node
"use strict";
const fs = require("node:fs");
const assert = require("node:assert/strict");
const client = fs.readFileSync("shared-static-views-client.js", "utf8");
const shared = fs.readFileSync("shared-static-views.js", "utf8");
const functions = fs.readFileSync("functions/shared-operational-views.js", "utf8");
const html = fs.readFileSync("index.html", "utf8");
assert.match(client, /subscribePersonale = \(\) => subscribeRegistry\("personale"\)/);
assert.match(client, /subscribeMezzi = \(\) => subscribeRegistry\("mezzi"\)/);
assert.match(client, /subscribeSquadre = function subscribeSquadreSafe/);
assert.match(client, /subscribeHoursStats = function subscribeHoursSafe/);
assert.match(shared, /callbacks = new Set/);
assert.match(functions, /registri__corrente/);
assert.match(functions, /squadre__\$\{date\}/);
assert.match(html, /shared-static-views-client\.js/);
console.log("✅ Avvio client instradato sulle viste condivise.");
console.log("✅ Listener deduplicati e trigger operativi presenti.");

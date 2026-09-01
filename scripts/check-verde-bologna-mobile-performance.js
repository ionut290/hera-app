#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("verde-bologna-parchi-mobile.js", "utf8");
const loader = fs.readFileSync("app-pure-utils.js", "utf8");
const html = fs.readFileSync("index.html", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");

assert.match(source, /const SEARCH_DEBOUNCE_MS = 180;/);
assert.match(source, /const LIST_RENDER_LIMIT = 60;/);
assert.match(source, /const LABEL_MARKER_LIMIT = 80;/);
assert.match(source, /fieldValue\(record, \["quartiere", "nomequartiere", "nome_quartiere"\]\)/);
assert.doesNotMatch(source, /fetchAllRecords\(QUARTIERI_DATASET_ID\)/);
assert.match(source, /state\.parks = parks\.map\(prepareRecord\)/);
assert.match(source, /state\.filtered\.slice\(0, LIST_RENDER_LIMIT\)/);
assert.match(source, /window\.L\.canvas\(\{ padding: 0\.5 \}\)/);
assert.match(source, /state\.map\.on\("zoomend", scheduleViewportRefresh\)/);
assert.match(source, /renderMapMarkers\(\{ fitMap: false \}\)/);
assert.match(source, /!fitMap && state\.map\.getZoom\(\) >= 15/);
assert.match(source, /state\.filterTimer = window\.setTimeout\(applyLiveFilter, SEARCH_DEBOUNCE_MS\)/);
assert.match(source, /removeLegacyQuartieriCache\(\)/);

const activation = source.match(/function activateParksMode\(\) \{[\s\S]*?\n  \}\n\n  function deactivateParksMode/);
assert.ok(activation, "Funzione activateParksMode non trovata");
assert.doesNotMatch(activation[0], /requestUserPosition\(\)/);

assert.match(loader, /verde-bologna-parchi-mobile\.js\?v=20260901-performance1/);
assert.match(html, /app-pure-utils\.js\?v=20260901-verde-performance1/);
assert.match(sw, /const CACHE_NAME = "varga-cantieri-shell-v202";/);
assert.match(sw, /app-pure-utils\.js\?v=20260901-verde-performance1/);

assert.doesNotMatch(source, /firebase|firestore|onSnapshot|addDoc|setDoc|updateDoc/);

console.log("Verde Bologna mobile performance checks passed.");

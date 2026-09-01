#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("verde-bologna-parchi-mobile.js", "utf8");
const baseSource = fs.readFileSync("verde-bologna.js", "utf8");
const operationalSource = fs.readFileSync("verde-bologna-operativo.js", "utf8");
const loader = fs.readFileSync("app-pure-utils.js", "utf8");
const html = fs.readFileSync("index.html", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");

assert.match(source, /const SEARCH_DEBOUNCE_MS = 180;/);
assert.match(source, /const FETCH_TIMEOUT_MS = 8000;/);
assert.match(source, /const LIST_RENDER_LIMIT = 60;/);
assert.match(source, /const LABEL_MARKER_LIMIT = 80;/);
assert.match(source, /fieldValue\(record, \["quartiere", "nomequartiere", "nome_quartiere"\]\)/);
assert.doesNotMatch(source, /fetchAllRecords\(QUARTIERI_DATASET_ID\)/);
assert.match(source, /varga-verde-bologna:\$\{PARKS_DATASET_ID\}:0:plain:/);
assert.match(source, /const controller = typeof AbortController === "function"/);
assert.match(source, /Promise\.allSettled\(remainingOffsets\.map/);
assert.match(source, /publishParkRecords\(firstPage\.records\)/);
assert.doesNotMatch(source, /while \(offset < total/);
assert.match(source, /state\.filtered\.slice\(0, LIST_RENDER_LIMIT\)/);
assert.match(source, /window\.L\.canvas\(\{ padding: 0\.5 \}\)/);
assert.match(source, /state\.map\.on\("zoomend", scheduleViewportRefresh\)/);
assert.match(source, /renderMapMarkers\(\{ fitMap: false \}\)/);
assert.match(source, /!fitMap && state\.map\.getZoom\(\) >= 15/);
assert.match(source, /state\.filterTimer = window\.setTimeout\(applyLiveFilter, SEARCH_DEBOUNCE_MS\)/);
assert.match(source, /removeLegacyQuartieriCache\(\)/);

const officialDatasets = [
  "un_gest", "alberi-manutenzioni", "popolazione-arborea", "siepi",
  "attrezzature_ludiche_ginniche_sportive", "arredo", "sgambatura_cani",
  "carta-tecnica-comunale-toponimi-parchi-e-giardini", "aree-verdi_entrate_centroidi",
  "aree-ortive", "verde_privato_urbanizzato"
];
officialDatasets.forEach((datasetId) => assert.match(baseSource, new RegExp(`id: "${datasetId.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"`)));
assert.match(baseSource, /const VIEWPORT_MAX_RECORDS = 500;/);
assert.match(baseSource, /const MOBILE_VIEWPORT_MAX_RECORDS = 140;/);
assert.match(baseSource, /const VIEWPORT_LIST_LIMIT = 60;/);
assert.match(baseSource, /const MOBILE_FULL_GEOMETRY_LIMIT = 40;/);
assert.match(baseSource, /const MOBILE_LABEL_MARKER_LIMIT = 80;/);
assert.match(baseSource, /within_distance\(\$\{geoField\}, geom'POINT\(/);
assert.match(baseSource, /state\.map\.on\("moveend zoomend", scheduleViewportLoad\)/);
assert.match(baseSource, /if \(zoom < 15\)/);
assert.match(baseSource, /const viewportLimit = mobileView\(\) \? MOBILE_VIEWPORT_MAX_RECORDS : VIEWPORT_MAX_RECORDS;/);
assert.match(baseSource, /state\.records\.slice\(0, VIEWPORT_LIST_LIMIT\)/);
assert.match(baseSource, /const lightweight = mobileView\(\) && state\.records\.length > MOBILE_FULL_GEOMETRY_LIMIT;/);
assert.match(baseSource, /classList\.toggle\("is-interactive", mobileView\(\) \|\| state\.fullscreen\)/);
assert.match(baseSource, /state\.map\.dragging\?\.enable\?\.\(\);/);
assert.match(baseSource, /state\.map\.touchZoom\?\.enable\?\.\(\);/);
assert.match(baseSource, /\.verde-bologna-map\.is-interactive\{touch-action:none!important\}/);
assert.doesNotMatch(baseSource, /state\.map\.dragging\?\.disable/);
assert.match(operationalSource, /\.verde-bologna-map\{width:100%!important;height:calc\(100dvh - 181px\)!important;min-height:420px!important;border-radius:0!important/);
assert.match(operationalSource, /\.verde-bologna-browser-head\{display:none!important\}/);
assert.match(operationalSource, /\.verde-bologna-map-toolbar\{position:absolute!important/);
assert.match(source, /\.verde-bologna-page\.vb-parks-advanced \.verde-bologna-parks-tools\{display:block\}/);
assert.match(baseSource, /data-vb-detail-index/);
assert.match(baseSource, /data-vb-toggle-details/);
assert.match(baseSource, /entryIndex >= 6/);
assert.match(baseSource, /NAVIGA VERSO L’ELEMENTO/);
assert.match(baseSource, /VISTA 360° E PERCORSO/);
assert.match(baseSource, /id="verde-bologna-map-style"/);
assert.match(baseSource, /<option value="classic">Classica<\/option>/);
assert.match(baseSource, /<option value="satellite">Satellite<\/option>/);
assert.match(baseSource, /<option value="hybrid">Ibrida<\/option>/);
assert.match(operationalSource, /#verde-bologna-map \.verde-bologna-marker-wrap/);
assert.match(operationalSource, /classList\.contains\("vb-parks-advanced"\)/);
assert.match(source, /\.vb-parks-advanced \.verde-bologna-marker-wrap\{display:none!important\}/);

const activation = source.match(/function activateParksMode\(\) \{[\s\S]*?\n  \}\n\n  function deactivateParksMode/);
assert.ok(activation, "Funzione activateParksMode non trovata");
assert.doesNotMatch(activation[0], /requestUserPosition\(\)/);

assert.match(loader, /verde-bologna\.js\?v=20260901-verde-compact1/);
assert.match(loader, /verde-bologna-operativo\.js\?v=20260901-verde-compact1/);
assert.match(loader, /verde-bologna-parchi-mobile\.js\?v=20260901-verde-compact1/);
assert.match(html, /app-pure-utils\.js\?v=20260901-verde-compact1/);
assert.match(html, /serviceWorker\.register\("\.\/sw\.js\?v=20260901-verde-compact1"/);
assert.match(sw, /const CACHE_NAME = "varga-cantieri-shell-v206";/);
assert.match(sw, /app-pure-utils\.js\?v=20260901-verde-compact1/);

assert.doesNotMatch(source, /firebase|firestore|onSnapshot|addDoc|setDoc|updateDoc/);
assert.doesNotMatch(baseSource, /firebase|firestore|onSnapshot|addDoc|setDoc|updateDoc/);

console.log("Verde Bologna mobile performance checks passed.");

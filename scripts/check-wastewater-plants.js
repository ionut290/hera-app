#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const html = fs.readFileSync("index.html", "utf8");
const source = fs.readFileSync("wastewater-plants.js", "utf8");
const css = fs.readFileSync("wastewater-plants.css", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");

assert.match(html, /id="open-wastewater-plants-btn"/);
assert.match(html, /id="wastewater-plants-page"/);
assert.match(html, /Censimento depuratori/);
assert.match(html, /Nome, codice, Comune, gestore o agglomerato/);
assert.match(html, /LA MIA POSIZIONE/);
assert.match(html, /VISTA 360° E PERCORSO/);
assert.match(html, /INVIA TRAMITE WHAZZUP/);
assert.match(source, /servizi-gis\.arpae\.it\/server\/rest\/services\/Geoportal\/ACQUEPressioni\/MapServer\/1\/query/);
assert.match(source, /resultOffset/);
assert.match(source, /exceededTransferLimit/);
assert.match(source, /resultRecordCount/);
assert.match(source, /COD_DEP/);
assert.match(source, /DEN_DEP/);
assert.match(source, /GESTORE/);
assert.match(source, /AE_PROG/);
assert.match(source, /COD_CIS/);
assert.match(source, /HeraStreetViewCards/);
assert.match(source, /Plugins\?\.HeraWhatsApp/);
assert.match(source, /whatsapp:\/\/send\?text=/);
assert.doesNotMatch(source, /wa\.me|web\.whatsapp\.com|api\.whatsapp\.com/);
assert.doesNotMatch(source, /firebase|firestore|collection\s*\(|onSnapshot|addDoc|setDoc|updateDoc/);
assert.match(css, /\.wastewater-plants-map-card--fullscreen/);
assert.match(css, /\.wastewater-plant-detail-section/);
assert.match(sw, /wastewater-plants\.js\?v=20260830a/);
assert.match(sw, /wastewater-plants\.css\?v=20260830a/);

console.log("Wastewater plants checks passed.");

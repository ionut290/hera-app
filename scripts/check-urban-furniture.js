#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const html = fs.readFileSync("index.html", "utf8");
const source = fs.readFileSync("urban-furniture.js", "utf8");
const css = fs.readFileSync("urban-furniture.css", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");

assert.match(html, /id="open-urban-furniture-btn"/);
assert.match(html, /id="urban-furniture-page"/);
assert.match(html, /CERCA NEL COMUNE/);
assert.match(html, /VICINO A ME/);
assert.match(html, /VISTA 360° E PERCORSO/);
assert.match(html, /INVIA TRAMITE WHAZZUP/);
assert.match(source, /https:\/\/overpass-api\.de\/api\/interpreter/);
assert.match(source, /area\["boundary"="administrative"\]\["admin_level"="8"\]/);
assert.match(source, /around:3000/);
assert.match(source, /amenity"="bench/);
assert.match(source, /amenity"="waste_basket/);
assert.match(source, /highway"="street_lamp/);
assert.match(source, /HeraStreetViewCards/);
assert.match(source, /Plugins\?\.HeraWhatsApp/);
assert.match(source, /whatsapp:\/\/send\?text=/);
assert.doesNotMatch(source, /wa\.me|web\.whatsapp\.com|api\.whatsapp\.com/);
assert.doesNotMatch(source, /firebase|firestore|collection\s*\(|onSnapshot|addDoc|setDoc|updateDoc/);
assert.match(css, /\.urban-furniture-map-card--fullscreen/);
assert.match(sw, /urban-furniture\.js\?v=20260830a/);
assert.match(sw, /urban-furniture\.css\?v=20260830a/);

console.log("Urban furniture checks passed.");

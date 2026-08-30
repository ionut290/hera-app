#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const proxyModule = require("../netlify/functions/urban-furniture.js");

const html = fs.readFileSync("index.html", "utf8");
const source = fs.readFileSync("urban-furniture.js", "utf8");
const proxy = fs.readFileSync("netlify/functions/urban-furniture.js", "utf8");
const css = fs.readFileSync("urban-furniture.css", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");

assert.match(html, /id="open-urban-furniture-btn"/);
assert.match(html, /id="urban-furniture-page"/);
assert.match(html, /CERCA NEL COMUNE/);
assert.match(html, /VICINO A ME/);
assert.match(html, /VISTA 360° E PERCORSO/);
assert.match(html, /INVIA TRAMITE WHAZZUP/);
assert.match(source, /\/api\/urban-furniture/);
assert.match(proxy, /overpass-api\.de\/api\/interpreter/);
assert.match(proxy, /overpass\.kumi\.systems\/api\/interpreter/);
assert.match(proxy, /area\["boundary"="administrative"\]\["admin_level"="8"\]/);
assert.match(proxy, /around:3000/);
assert.match(proxy, /amenity"="bench/);
assert.match(proxy, /amenity"="waste_basket/);
assert.match(proxy, /highway"="street_lamp/);
assert.match(proxy, /leisure"="playground/);
assert.match(proxy, /emergency"="fire_hydrant/);
assert.match(proxy, /tourism"="artwork/);
assert.match(source, /HeraStreetViewCards/);
assert.match(source, /Plugins\?\.HeraWhatsApp/);
assert.match(source, /whatsapp:\/\/send\?text=/);
assert.doesNotMatch(source, /wa\.me|web\.whatsapp\.com|api\.whatsapp\.com/);
assert.doesNotMatch(source, /firebase|firestore|collection\s*\(|onSnapshot|addDoc|setDoc|updateDoc/);
assert.doesNotMatch(proxy, /firebase|firestore|collection\s*\(|onSnapshot|addDoc|setDoc|updateDoc/);
assert.match(css, /\.urban-furniture-map-card--fullscreen/);
assert.match(sw, /urban-furniture\.js\?v=20260830-allcategories1/);
assert.match(sw, /urban-furniture\.css\?v=20260830-allcategories1/);
assert.equal(typeof proxyModule.handler, "function");
const categories = Object.keys(proxyModule._test.CATEGORIES);
assert.equal(categories.length, 39);
for (const category of categories) {
  assert.match(html, new RegExp(`<option value=["']${category}["']`));
  assert.match(source, new RegExp(`\\b${category}:`));
  assert.doesNotThrow(() => proxyModule._test.buildQuery({ mode: "municipality", municipality: "Bologna", category }));
}
assert.match(proxyModule._test.buildQuery({ mode: "municipality", municipality: "Bologna", category: "bench" }), /area\.municipality/);
assert.match(proxyModule._test.buildQuery({ mode: "nearby", lat: "44.49", lon: "11.34", category: "all" }), /around:3000/);

console.log("Urban furniture checks passed.");

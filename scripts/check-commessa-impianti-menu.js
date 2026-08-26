#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const html = fs.readFileSync("index.html", "utf8");
const feature = fs.readFileSync("commessa-impianti-menu.js", "utf8");
const accounting = fs.readFileSync("accounting-v2.js", "utf8");
const css = fs.readFileSync("style.css", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");

assert.match(html, /id="commessa-plants-menu-btn"/);
for (const action of ["add", "edit", "import", "export", "prices", "advanced"]) {
  assert.match(feature, new RegExp(`data-commessa-mobile-action="${action}"`));
}
assert.match(feature, /canManageData\(\)/);
assert.match(feature, /openManagementPanel\("commesse"\)/);
assert.match(feature, /AccountingV2\.openMobileHub\(commessa\)/);
assert.match(feature, /id="commessa-mobile-management"/);
assert.match(accounting, /mobileGroups/);
assert.match(accounting, /Dati impianto/);
assert.match(accounting, /Lavorazione/);
assert.match(accounting, /Esecuzione/);
assert.match(accounting, /async function createMobilePlant/);
assert.match(accounting, /async function saveMobilePlant/);
assert.match(css, /\.commessa-dashboard-head \.commessa-plants-menu-wrap\s*{[^}]*position:\s*absolute/s);
assert.match(css, /\.commessa-mobile-plant-form/);
assert.doesNotMatch(feature, /\bdb\.|\bfirebase\.|\.collection\(|\.onSnapshot\(/);
assert.match(serviceWorker, /commessa-impianti-menu\.js\?v=20260826b/);
assert.match(serviceWorker, /accounting-v2\.js\?v=20260826-commessa-mobile1/);
assert.match(serviceWorker, /style\.css\?v=20260826-commessa-mobile1/);

const accountingIndex = html.indexOf("accounting-v2.js");
const featureIndex = html.indexOf("commessa-impianti-menu.js");
const fattoIndex = html.indexOf("fatto-button-immediate.js");
assert.ok(accountingIndex >= 0 && featureIndex > accountingIndex, "Il menu deve riusare la gestione impianti già caricata.");
assert.ok(fattoIndex > featureIndex, "Il nuovo menu deve restare esterno al componente FATTO sigillato.");

console.log("✅ Gestione commessa mobile collegata ai dati esistenti con pulsante sovrapposto alla testata.");

#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const html = fs.readFileSync("index.html", "utf8");
const feature = fs.readFileSync("commessa-impianti-menu.js", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");

assert.match(html, /id="commessa-plants-menu-btn"/);
for (const action of ["add", "edit", "import", "export", "prices"]) {
  assert.match(html, new RegExp(`data-commessa-plants-action="${action}"`));
}
assert.match(feature, /canManageData\(\)/);
assert.match(feature, /openManagementPanel\("commesse"\)/);
assert.match(feature, /openImpiantiManagement\(commessa\)/);
assert.match(feature, /add-management-impianto-btn/);
assert.match(feature, /impianti-import-card/);
assert.match(feature, /export-all-impianti-btn/);
assert.match(feature, /open-prezziario-btn/);
assert.doesNotMatch(feature, /\bdb\.|\bfirebase\.|\.collection\(|\.onSnapshot\(/);
assert.match(serviceWorker, /commessa-impianti-menu\.js\?v=20260826a/);

const accountingIndex = html.indexOf("accounting-v2.js");
const featureIndex = html.indexOf("commessa-impianti-menu.js");
const fattoIndex = html.indexOf("fatto-button-immediate.js");
assert.ok(accountingIndex >= 0 && featureIndex > accountingIndex, "Il menu deve riusare la gestione impianti già caricata.");
assert.ok(fattoIndex > featureIndex, "Il nuovo menu deve restare esterno al componente FATTO sigillato.");

console.log("✅ Menu Gestione impianti collegato alla schermata esistente senza nuove operazioni Firestore.");

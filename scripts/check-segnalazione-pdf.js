"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const html = fs.readFileSync("index.html", "utf8");

assert.match(html, /id="segnalazione-genera-pdf-btn"[^>]*type="submit"/);
assert.match(html, /jspdf(?:\.umd)?\.min\.js/);
assert.match(app, /async function generateSegnalazionePdf\(event\)/);
assert.match(app, /new jsPDF\(\{ orientation: "portrait", unit: "mm", format: "a4" \}\)/);
assert.match(app, /lastSegnalazionePdfBlob = doc\.output\("blob"\)/);
assert.match(app, /lastSegnalazionePdfName = `scheda-segnalazione-/);
assert.match(app, /doc\.save\(lastSegnalazionePdfName\)/);
assert.match(app, /type: "application\/pdf"/);

console.log("Segnalazione PDF checks passed.");

#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const style = fs.readFileSync("style.css", "utf8");

const requiredAppMarkers = [
  "const whazzupPhotoNotesByImpianto = new Map();",
  "function normalizeWhazzupPhotoNotes(notes, photoCount)",
  "notes: normalizeWhazzupPhotoNotes(notes, files.length)",
  "function getWhazzupPhotoNotes(impianto)",
  "async function saveWhazzupPhotoNotes(impianto, notes)",
  'data-photo-note-index="${index}"',
  'maxlength="180"',
  "async function addWhazzupNoteToPhoto(file, note)",
  'context.fillText("NOTA FOTO"',
  "async function buildWhazzupShareFilesWithNotes(impianto, files)",
  "const orderedFiles = await buildWhazzupShareFilesWithNotes(impianto, files);"
];

requiredAppMarkers.forEach((marker) => {
  assert.ok(app.includes(marker), `Marker note foto mancante: ${marker}`);
});

assert.match(app, /whazzupPhotoNotesByImpianto\.delete\(key\)/);
assert.match(app, /saveWhazzupPhotoSelection\(impianto, remaining, remainingNotes\)/);
assert.match(app, /await flushVisibleNotes\(\);[\s\S]*close\(\);[\s\S]*handleCompletedImpiantoWhatsAppClick/);
assert.match(app, /if \(!normalizedNote\) return file;/);
assert.match(app, /return buildOrderedWhazzupShareFiles\(annotatedFiles\);/);

assert.match(style, /\.whazzup-photo-note\s*\{/);
assert.match(style, /\.whazzup-photo-note textarea\s*\{/);
assert.match(style, /\.whazzup-photo-note textarea:focus\s*\{/);

console.log("Whazzup per-photo note checks passed.");

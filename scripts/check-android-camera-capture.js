#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const style = fs.readFileSync("style.css", "utf8");
const manifest = fs.readFileSync("android/app/src/main/AndroidManifest.xml", "utf8");

const requiredAppMarkers = [
  'const useCamera = options.source === "camera";',
  'input.multiple = !useCamera && options.mode !== "replace-one";',
  'if (useCamera) input.setAttribute("capture", "environment");',
  "function openWhazzupPhotoSourceChooser(impianto, button, options = {})",
  'data-photo-source="camera"',
  'data-photo-source="gallery"',
  "Scatta foto",
  "Scegli dalla galleria",
  'pickWhazzupPhotos(impianto, button, { ...options, source });',
  'openWhazzupPhotoSourceChooser(impianto, button, { mode: "replace-all" });',
  'openWhazzupPhotoSourceChooser(impianto, button, { mode: "replace-one", index });',
  'else openWhazzupPhotoSourceChooser(impianto, button, { mode: "replace-all" });'
];

requiredAppMarkers.forEach((marker) => {
  assert.ok(app.includes(marker), `Marker fotocamera Android mancante: ${marker}`);
});

assert.match(style, /\.whazzup-photo-source-card\s*\{/);
assert.match(style, /\.whazzup-photo-source-actions \.whazzup-photo-camera-btn\s*\{/);
assert.match(style, /\.whazzup-photo-source-chooser\s*\{\s*align-items:\s*end;/);

const hasNativeContinuousCamera = manifest.includes('.camera.HeraContinuousCameraActivity');
const cameraPermission = /<uses-permission\s+android:name="android\.permission\.CAMERA"/;
if (hasNativeContinuousCamera) {
  assert.match(
    manifest,
    cameraPermission,
    "La fotocamera continua Android nativa richiede android.permission.CAMERA"
  );
} else {
  assert.doesNotMatch(
    manifest,
    cameraPermission,
    "Il vecchio flusso IMAGE_CAPTURE esterno non deve richiedere android.permission.CAMERA"
  );
}

assert.match(manifest, /<queries>[\s\S]*<action android:name="android\.media\.action\.IMAGE_CAPTURE"\s*\/>[\s\S]*<\/queries>/);
const captureStart = app.indexOf("function pickWhazzupPhotos");
const captureEnd = app.indexOf("function openWhazzupPhotoManager", captureStart);
assert.ok(captureStart >= 0 && captureEnd > captureStart, "Blocco selezione foto non trovato");
assert.doesNotMatch(app.slice(captureStart, captureEnd), /navigator\.mediaDevices\.getUserMedia/);

console.log("Android direct/continuous camera capture checks passed.");

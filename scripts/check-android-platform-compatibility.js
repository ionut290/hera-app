#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const read = (path) => fs.readFileSync(path, "utf8");
const mainActivity = read("android/app/src/main/java/it/vargacantieri/hera/MainActivity.java");
const cameraActivity = read("android/app/src/main/java/it/vargacantieri/hera/camera/HeraContinuousCameraActivity.java");
const manifest = read("android/app/src/main/AndroidManifest.xml");
const androidDependencies = read("android/app/capacitor.build.gradle");
const html = read("index.html");
const workflow = read(".github/workflows/build-android-aab.yml");

function assertEdgeToEdgeBeforeSuper(source, label) {
  assert.match(source, /import androidx\.activity\.EdgeToEdge;/, `${label}: import EdgeToEdge mancante.`);
  const enablePosition = source.indexOf("EdgeToEdge.enable(this);");
  const superPosition = source.indexOf("super.onCreate(savedInstanceState);");
  assert.ok(enablePosition >= 0 && superPosition > enablePosition, `${label}: edge-to-edge deve essere abilitato prima di super.onCreate.`);
}

assertEdgeToEdgeBeforeSuper(mainActivity, "MainActivity");
assertEdgeToEdgeBeforeSuper(cameraActivity, "Fotocamera continua");

// Navigation bars overlap all buttons when an edge-to-edge Capacitor WebView
// fills the display. The protection belongs to the native content viewport,
// not a small set of CSS selectors. Guard all critical parts of this contract.
for (const required of [
  "protectWebContentFromAndroidSystemNavigation();",
  "findViewById(android.R.id.content)",
  "ViewCompat.setOnApplyWindowInsetsListener(contentRoot",
  "WindowInsetsCompat.Type.systemBars()",
  "WindowInsetsCompat.Type.displayCutout()",
  "windowInsets.getInsets(safeTypes)",
  "view.setPadding(left, top, right, bottom)",
  "new WindowInsetsCompat.Builder(windowInsets)",
  ".setInsets(safeTypes, Insets.NONE)",
  "ViewCompat.requestApplyInsets(contentRoot)"
]) {
  assert.ok(mainActivity.includes(required),
    "Safe area Android non completa, pulsanti a rischio sotto navigazione: " + required);
}
assert.match(mainActivity, /initialBottom \+ safe\.bottom/,
  "Il margine di navigazione deve sommarsi al padding originale, senza sovrascriverlo.");
assert.match(mainActivity, /initialLeft \+ safe\.left/,
  "Le barre laterali in landscape devono proteggere anche il bordo sinistro.");
assert.match(mainActivity, /initialRight \+ safe\.right/,
  "Le barre laterali in landscape devono proteggere anche il bordo destro.");
assert.match(mainActivity, /WindowCompat\.setDecorFitsSystemWindows\(getWindow\(\), false\)/,
  "Il listener deve ricevere gli insets completi in modalità edge-to-edge.");
assert.doesNotMatch(mainActivity, /hide\(WindowInsetsCompat\.Type\.navigationBars\(\)\)/,
  "La correzione non deve nascondere i pulsanti di navigazione Android.");

assert.doesNotMatch(cameraActivity, /setStatusBarColor|setNavigationBarColor/, "La fotocamera usa ancora API colore deprecate.");
for (const required of [
  "WindowCompat.getInsetsController",
  "ViewCompat.setOnApplyWindowInsetsListener",
  "WindowInsetsCompat.Type.systemBars()",
  "WindowInsetsCompat.Type.displayCutout()",
  "ViewCompat.requestApplyInsets(root)"
]) {
  assert.ok(cameraActivity.includes(required), `Gestione inset fotocamera incompleta: ${required}`);
}

assert.doesNotMatch(manifest, /android:screenOrientation=/, "Il manifest limita ancora l’orientamento su schermi grandi.");
assert.match(html, /<meta name="viewport" content="[^"]*viewport-fit=cover[^"]*">/, "La WebView non espone le safe area edge-to-edge.");
assert.match(androidDependencies, /def cameraxVersion = "1\.4\.2"/, "CameraX non è sulla versione verificata per pagine da 16 kB.");

for (const required of [
  "npm run check:android-platform",
  "Verify 16 KB native library alignment",
  "readelf -lW",
  "alignment < 16384"
]) {
  assert.ok(workflow.includes(required), `Controllo AAB 16 kB incompleto: ${required}`);
}

console.log("Compatibilità Android verificata: edge-to-edge, schermi grandi e librerie native da 16 kB protetti.");

#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

function read(path) {
  assert.ok(fs.existsSync(path), `File mancante: ${path}`);
  return fs.readFileSync(path, "utf8");
}

const app = read("app.js");
const index = read("index.html");
const authFix = read("auth-login-fix.js");
const updateFeature = read("update-app-feature.js");
const updateFeatureExecutable = updateFeature
  .replace(/\/\*[\s\S]*?\*\//g, "")
  .replace(/^\s*\/\/.*$/gm, "");
const netlifyHeaders = read("_headers");
const approval = read("approval-access.js");
const activeCommesseGuard = read("active-commesse-first-boot-guard.js");
const squadraSync = read("squadra-current-save-sync.js");
const doneImmediate = read("fatto-button-immediate.js");
const hoursGuard = read("hours-source-explicit-guard.js");
const sharedViews = read("shared-static-views-client.js");
const androidOrder = read("android-whazzup-photo-order.js");
const serviceWorker = read("sw.js");
const headerRuntime = read("header-menu-runtime.js");

const flows = [];
function flow(name, checks) {
  checks();
  flows.push(name);
}

flow("1. Accesso stabile", () => {
  assert.match(app, /ensureAuthLocalPersistence\(/);
  assert.match(app, /PERSISTED_SESSION_KEY/);
  assert.match(authFix, /installProfileAccessGuard\(/);
  assert.match(authFix, /installAuthStartupController/);
  assert.match(authFix, /Auth\?\.Persistence\?\.LOCAL|Auth\.Persistence\.LOCAL/);
  assert.match(authFix, /onIdTokenChanged/);
  assert.match(authFix, /MutationObserver/);
  assert.match(authFix, /HeraAuthStartupController/);
  assert.doesNotMatch(headerRuntime, /loadAutoLoginFeature/);
  assert.doesNotMatch(headerRuntime, /auto-login-saved-credentials\.js/);
  assert.doesNotMatch(headerRuntime, /header-menu-runtime-original\.js/);
  assert.doesNotMatch(headerRuntime, /commessa-listener-cleanup\.js/);
  assert.match(headerRuntime, /__heraCommessaListenerCleanupInstalled/);
  assert.doesNotMatch(serviceWorker, /auto-login-saved-credentials\.js/);
  assert.doesNotMatch(serviceWorker, /header-menu-runtime-original\.js/);
  assert.match(updateFeature, /APP_CACHE_PREFIXES/);
  assert.match(updateFeature, /cacheNames\.filter\(isAppShellCache\)/);
  assert.doesNotMatch(updateFeatureExecutable, /firebase\.auth\(\)\.signOut\(/);
  assert.doesNotMatch(updateFeatureExecutable, /localStorage\.clear\(/);
  assert.doesNotMatch(updateFeatureExecutable, /sessionStorage\.clear\(/);
  assert.doesNotMatch(updateFeatureExecutable, /indexedDB\.deleteDatabase\(/);
  for (const path of ["/", "/index.html", "/sw.js", "/auth-login-fix.js"]) {
    assert.match(netlifyHeaders, new RegExp(`${path.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\n\\s+Cache-Control: no-cache, must-revalidate`));
  }
  assert.match(approval, /window\.HeraAccessApproval\s*=\s*\{/);
});

flow("2. Lavori di oggi / commesse e squadre", () => {
  assert.match(activeCommesseGuard, /HeraActiveCommesseFirstBootGuard/);
  assert.match(activeCommesseGuard, /installIndexReadGuard\(/);
  assert.match(activeCommesseGuard, /installListenerGuard\(/);
  assert.match(activeCommesseGuard, /failOpenRootListeners/);
  assert.match(squadraSync, /squadr/i);
  assert.match(app, /renderTodaySummary\(/);
});

flow("3. Ricerca impianti", () => {
  assert.match(index, /type="search"/i);
  assert.match(app, /impiant/i);
  assert.match(app, /search|cerca/i);
});

flow("4. NAVIGA - graffetta - FATTO", () => {
  assert.match(app, /NAVIGA|navigate|navigation/i);
  assert.match(doneImmediate, /FATTO|fatto/i);
  assert.match(index, /📎|graffetta|foto/i);
});

flow("5. Whazzup Android: messaggio prima delle foto", () => {
  const messageIndex = androidOrder.indexOf("safeOpenWhatsAppMessage(message)");
  const dedicatedPhotoIndex = androidOrder.indexOf("await sharePhotosThroughDedicatedPlugin(dedicatedPlugin, orderedFiles)");
  const fallbackPhotoIndex = androidOrder.indexOf("await plugins.share.share");
  assert.ok(messageIndex >= 0, "Apertura messaggio Whazzup mancante");
  assert.ok(dedicatedPhotoIndex >= 0 && messageIndex < dedicatedPhotoIndex, "Il messaggio deve precedere le foto nel plugin dedicato");
  assert.ok(fallbackPhotoIndex >= 0 && messageIndex < fallbackPhotoIndex, "Il messaggio deve precedere le foto nel fallback Android");
});

flow("6. Gestione ore", () => {
  assert.match(app, /hoursReport/);
  assert.match(hoursGuard, /explicitHoursSourceGuard/);
  assert.match(sharedViews, /SAFE CALENDAR GUARD/);
});

flow("7. Segnalazioni e note", () => {
  assert.match(app, /commessaNote/);
  assert.match(app, /segnalazion/i);
  assert.match(app, /enqueueOfflineMutation\("commessaNote"/);
});

flow("8. Sicurezza e documenti consultabili", () => {
  assert.match(index, /POS - Documenti sicurezza|open-pos-btn/);
  assert.match(serviceWorker, /index\.html|app\.js/);
  assert.match(serviceWorker, /cache/i);
});

flow("9. Offline e sincronizzazione automatica", () => {
  assert.match(app, /PENDING_OFFLINE_MUTATIONS_KEY/);
  assert.match(app, /function enqueueOfflineMutation\(/);
  assert.match(app, /async function syncPendingOfflineMutations\(/);
  assert.match(app, /window\.addEventListener\("online"[\s\S]*syncPendingOfflineMutations\(\)/);
  assert.match(app, /enqueueOfflineMutation\("hoursReport"/);
});

console.log(`Critical flow checks passed (${flows.length}/9):`);
for (const name of flows) console.log(` - ${name}`);

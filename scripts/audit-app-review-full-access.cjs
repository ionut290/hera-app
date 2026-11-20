"use strict";

/**
 * App Review 2.1 full-function access READINESS CHECK.
 * Read-only static audit of tracked source files; no user account access,
 * no Firestore queries, no credentials and NO production changes.
 *
 *   node scripts/audit-app-review-full-access.cjs
 *   node scripts/audit-app-review-full-access.cjs --require-ready
 *
 * When the check is not ready, the default mode reports the blockers and
 * exits normally so it is visible in CI. --require-ready fails closed.
 */
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const root = path.resolve(__dirname, "..");
const read = relative => fs.readFileSync(path.join(root, relative), "utf8");

const coverage = [
  ["Accesso e profilo", "account-privacy-controls.js"],
  ["Commesse e gestione", "management-v2.js"],
  ["Impianti e navigazione", "commessa-impianti-menu.js"],
  ["Timbrature e promemoria", "timbratura-reminders.js"],
  ["Squadre e assegnazioni", "squadra-current-save-sync.js"],
  ["Documenti", "documents.js"],
  ["Calendario", "administrative-calendar.js"],
  ["Contabilità e consuntivi", "accounting-v2.js"]
].map(([feature, file]) => ({
  feature, file, sourcePresent: fs.existsSync(path.join(root, file)) &&
    fs.statSync(path.join(root, file)).size > 0,
  // IMPORTANT: code presence does not prove functional/demo access!
  demoFunctionalTestPassed: false
}));

const firestore = read("firestore.rules");
const storage = read("storage.rules");
const privacyClient = read("account-privacy-controls.js");
const demoUI = read("apple-review-demo-deletion-ui.js");
const packager = read("scripts/prepare-capacitor-web.js");

// Missing code is a release blocker, not an unstructured crash: operators
// need the complete reasons even when one module disappears in a refactor.
const missingFeatures = coverage.filter(row => !row.sourcePresent);
assert.match(firestore, /function isAppleReviewAccount\(\)/);
assert.match(storage, /function isAppleReviewAccount\(\)/);
assert.match(packager, /excludedFilePrefixes\s*=\s*\["apple-review-"\]/);

const blockers = [];
for (const item of missingFeatures) {
  blockers.push("MISSING_FEATURE_SOURCE: " + item.file + " (" + item.feature + ")");
}
if (/&& !isAppleReviewAccount\(\)/.test(firestore)) {
  blockers.push("REVIEW_CLAIM_BLOCKS_OPERATIONAL_FIRESTORE: Apple reviewer claim is excluded from signedIn(); commesse/personale/timbrature are not reviewable.");
}
if (/!isAppleReviewAccount\(\)/.test(storage)) {
  blockers.push("REVIEW_CLAIM_BLOCKS_STORAGE: reviewer cannot access normal document storage; demo must have synthetic separate document storage.");
}
if (privacyClient.includes("window.__VARGA_ACCOUNT_DELETION_ENABLED === true")) {
  blockers.push("DELETION_UI_DEFAULT_OFF: ordinary iOS build hides self-service deletion pending privacy signoff.");
}
if (/apple-review-demo-deletion-ui\.js/.test(packager) || demoUI.includes("internal demo build")) {
  blockers.push("DEMO_DELETE_UI_SEPARATE: internal demo-only preview is not shipped in standard native app.");
}
blockers.push("NO_VERIFIED_FULL_FEATURE_SANDBOX: Firebase emulator fixtures are not a deployed reviewer sandbox; no iPhone feature-by-feature access proof is available.");
// Last observed via read-only Firebase console audit on 20 Nov 2026:
// the demo identity EXISTED but was DISABLED. This is a dated observation,
// NOT a live status check. Do not claim it was deleted or is now enabled.
blockers.push("NO_ACTIVE_VERIFIED_REVIEW_LOGIN: demo account was present but DISABLED on 2026-11-20; current active login and isolation are unverified.");
blockers.push("NO_CONFIRMED_GENERAL_ACCOUNT_DELETION: ordinary account-deletion UI remains gated and real privacy fulfillment has not been approved/deployed.");
blockers.push("NO_APPLE_SIGNIN_DELETION_DEVICE_TEST: iPhone Apple login works per user report, but Apple credential revocation and ordinary account deletion are NOT verified.");

const report = {
  title: "VARGA CANTIERI - Apple App Review complete-access readiness",
  status: blockers.length ? "BLOCKED" : "READY",
  evidence: {
    appleTestFlightBuildLastConfirmed: "1.0 (20015)",
    liveFirebaseAuthCheckedByThisScript: false,
    realIphoneAccountDeletionPassed: false,
    fullReviewerFeatureAccessPassed: false
  },
  safety: "Never unlock production Firestore/Storage reviewer access or expose employee data to make audit green.",
  coverage,
  blockers
};
console.log(JSON.stringify(report, null, 2));
if (process.argv.includes("--require-ready") && blockers.length) process.exitCode = 1;

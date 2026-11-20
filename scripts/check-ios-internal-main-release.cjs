"use strict";
// Offline integrity check; no credentials, build signing or Apple network.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const internal = fs.readFileSync(".github/workflows/ios-internal-testflight-main.yml","utf8");
const publicApp = fs.readFileSync(".github/workflows/ios-appstore-candidate.yml","utf8");
const ordinary = fs.readFileSync(".github/workflows/deploy-firebase-functions.yml","utf8");
const must = (s,fragment,message) =>
  assert.ok(s.includes(fragment), message + ": " + fragment);

assert.match(internal,/^on:\n  workflow_dispatch:/m);
assert.doesNotMatch(internal,/^  push:|^  pull_request:|^  schedule:/m,
  "Only manual TestFlight uploads are permitted");
assert.match(internal,/default: DO_NOT_UPLOAD/);
must(internal,"UPLOAD_INTERNAL_TESTFLIGHT_ONLY","Explicit internal approval required");
assert.doesNotMatch(internal,/UPLOAD_APP_STORE_CANDIDATE_ONLY/,
  "Never dispatch the public App Store upload from this workflow");
for (const value of [
  'refs/heads/main',
  'ionut290/hera-app',
  'GITHUB_ACTOR',
  'GITHUB_TRIGGERING_ACTOR',
  'GITHUB_RUN_ATTEMPT',
  'Reviewed SHA must match the exact main commit',
  'bash scripts/check-ios-appstore-build-number.sh "$IOS_BUILD_NUMBER"',
  'environment: ios-internal-testflight',
  'needs: preflight',
  'test -s ios/App/App/App.entitlements',
  'ruby scripts/install-ios-google-signin.rb',
  'ruby scripts/install-ios-apple-signin.rb',
  'python3 scripts/configure-ios-firebase.py',
  'node scripts/check-account-privacy-client-runtime.js',
  'node scripts/check-app-review-readiness-failclosed.cjs',
  'Not submitted to App Review'
]) {
  if (value === 'Not submitted to App Review') {
    assert.match(internal,/NOT submitted to App Review/);
  } else {
    must(internal,value,"Missing private TestFlight safety prerequisite");
  }
}
assert.ok(internal.split("GITHUB_RUN_ATTEMPT").length>=3,
  "Both preflight and upload must reject job reruns independently");
assert.match(internal,/node scripts\/audit-app-review-full-access\.cjs > /,
  "Must document actual review readiness for internal tester");
assert.doesNotMatch(internal,/node scripts\/audit-app-review-full-access\.cjs --require-ready/,
  "Unmet public-review requirements must not prevent internal testing");
assert.match(publicApp,/node scripts\/audit-app-review-full-access\.cjs --require-ready/,
  "Never weaken the PUBLIC App Review upload gate");
assert.match(internal,/"method": "app-store-connect"/,
  "Internal build must use signed ASC upload; no public review submission");
assert.match(internal,/"destination": "upload"/);
assert.doesNotMatch(internal,/submit_for_review|submitForReview|--submit-for-review|deliver.*submit/i);
for (const forbidden of [
  "firebase deploy", "firestore:rules", ":setIamPolicy",
  "revokeRefreshTokens(", "auth().deleteUser(",
  "__VARGA_ACCOUNT_DELETION_REQUEST_ENABLED = true",
  "__VARGA_ACCOUNT_DELETION_ENABLED = true",
  "deleteOwnVargaAccount("
]) {
  assert.ok(!internal.includes(forbidden),
    "Internal build is not allowed to mutate real Firebase: " + forbidden);
}
assert.match(ordinary,/default: PREVIEW_ONLY/);
assert.doesNotMatch(ordinary,/^  push:/m);
const secretPosition = internal.indexOf('secrets.IOS_APPSTORE_PRIVATE_KEY_BASE64');
const uploadPosition = internal.indexOf('\n  upload-internal-testflight:');
assert.ok(secretPosition>uploadPosition && uploadPosition>=0,
  "Apple private key must not be exposed to the non-secret preflight job");
assert.match(internal,/persist-credentials: false/);
assert.match(internal,/if: always\(\)/,
  "Signing material must be removed after upload failure/success");
console.log("PASS: internal TestFlight only, owner+SHA+fresh confirmation, App Review separately blocked, no Firebase mutation or public submission.");

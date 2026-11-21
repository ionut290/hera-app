"use strict";

// Offline, no Apple API credentials, no Firebase access and no uploads.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const cp = require("node:child_process");
const script = "scripts/check-ios-appstore-build-number.sh";
const workflow = fs.readFileSync(".github/workflows/ios-appstore-candidate.yml", "utf8");
const internal = fs.readFileSync(".github/workflows/ios-internal-testflight-main.yml", "utf8");
assert.match(internal, /default: "20017"/);
assert.match(internal, /TestFlight build 1\.0 \(20016\)/);
assert.equal((internal.match(/bash scripts\/check-ios-appstore-build-number\.sh "\$IOS_BUILD_NUMBER"/g) || []).length, 2,
  "Both stages of internal TestFlight must reject build reuse");
assert.match(workflow, /default: "20017"/);
assert.match(workflow, /TestFlight build 1\.0 \(20016\)/);
const uses = workflow.match(/bash scripts\/check-ios-appstore-build-number\.sh "\$IOS_BUILD_NUMBER"/g) || [];
assert.equal(uses.length, 2, "Both Linux preflight and macOS signing must enforce the same build guard");
assert.ok(workflow.includes("node scripts/audit-app-review-full-access.cjs --require-ready"),
  "App Store upload must stay blocked without actual App Review clearance");
assert.ok(workflow.includes("needs: preflight"));

const preflight = workflow.split("\n  preflight:")[1]?.split("\n  upload-appstore-candidate:")[0];
assert.ok(preflight, "Missing preflight job");
const checkoutPos = preflight.indexOf("- name: Checkout release readiness assessment");
const inputPos = preflight.indexOf("- name: Validate manual input");
const scriptPos = preflight.indexOf('bash scripts/check-ios-appstore-build-number.sh "$IOS_BUILD_NUMBER"');
const readinessPos = preflight.indexOf("node scripts/audit-app-review-full-access.cjs --require-ready");
assert.ok(checkoutPos !== -1 && checkoutPos < inputPos && inputPos < scriptPos && scriptPos < readinessPos,
  "Linux preflight must checkout files before executing repo scripts, without weakening the release gate");
assert.match(preflight.slice(checkoutPos, inputPos), /uses: actions\/checkout@v4/);
assert.equal((preflight.match(/uses: actions\/checkout@v4/g) || []).length, 1,
  "Only one preflight checkout is needed; no redundant checkout");

assert.match(workflow, /202704/);
assert.match(workflow, /SDK_MAJOR.*-lt 27/);
const run = number => cp.spawnSync("bash", [script, number], {encoding: "utf8"});
for (const number of ["20017", "20018", "22000", "999999"]) {
  assert.equal(run(number).status, 0, "Valid future build rejected: " + number);
}
for (const number of ["10004", "20010", "20011", "20012", "20013", "20014",
  "20015", "20016", "020017", "20017x", "20017.1", "0", "-1", "", "1000000", "2e4",
  "20016 "]) {
  assert.notEqual(run(number).status, 0, "Old/invalid build accepted: " + JSON.stringify(number));
}
console.log("PASS: both App Store and TestFlight stages reject reused iOS build numbers (<=20016), preserve review gate and future iOS 27 SDK guard.");

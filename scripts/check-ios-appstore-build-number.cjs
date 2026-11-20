"use strict";

// Offline, no Apple API credentials, no Firebase access and no uploads.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const cp = require("node:child_process");
const script = "scripts/check-ios-appstore-build-number.sh";
const workflow = fs.readFileSync(".github/workflows/ios-appstore-candidate.yml", "utf8");
assert.match(workflow, /default: "20012"/);
assert.match(workflow, /TestFlight build 1\.0 \(20011\)/);
const uses = workflow.match(/bash scripts\/check-ios-appstore-build-number\.sh "\$IOS_BUILD_NUMBER"/g) || [];
assert.equal(uses.length, 2, "Both Linux preflight and macOS signing must enforce the same build guard");
assert.ok(workflow.includes("node scripts/audit-app-review-full-access.cjs --require-ready"),
  "App Store upload must stay blocked without actual App Review clearance");
assert.ok(workflow.includes("needs: preflight"));
assert.match(workflow, /202704/);
assert.match(workflow, /SDK_MAJOR.*-lt 27/);
const run = number => cp.spawnSync("bash", [script, number], {encoding: "utf8"});
for (const number of ["20012", "20013", "22000", "999999"]) {
  assert.equal(run(number).status, 0, "Valid future build rejected: " + number);
}
for (const number of ["10004", "20010", "20011", "020012", "20012x", "20012.1", "0",
  "-1", "", "1000000", "2e4", "20011 "]) {
  assert.notEqual(run(number).status, 0, "Old/invalid build accepted: " + JSON.stringify(number));
}
console.log("PASS: both manual App Store upload stages reject obsolete build numbers (<=20011), preserve review gate and future iOS 27 SDK guard.");

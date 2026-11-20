"use strict";
// Offline release-gate regression: no Firebase, Apple, GitHub API or secrets.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const cp = require("node:child_process");
const vm = require("node:vm");

const program = path.resolve("scripts/audit-app-review-full-access.cjs");
const readReport = args => {
  const run = cp.spawnSync(process.execPath, [program, ...args], {
    encoding: "utf8", timeout: 10000
  });
  assert.ifError(run.error);
  return { code: run.status, data: JSON.parse(run.stdout), stderr: run.stderr };
};

const observation = readReport([]);
assert.equal(observation.code, 0, "Read-only reporting must not throw");
assert.equal(observation.data.status, "BLOCKED");
assert.equal(observation.data.evidence.appleTestFlightBuildLastConfirmed, "1.0 (20015)");
assert.equal(observation.data.evidence.liveFirebaseAuthCheckedByThisScript, false);
assert.equal(observation.data.evidence.realIphoneAccountDeletionPassed, false);
assert.equal(observation.data.evidence.fullReviewerFeatureAccessPassed, false);
assert.ok(observation.data.coverage.length >= 8);
assert.ok(observation.data.coverage.every(row =>
  typeof row.sourcePresent === "boolean" &&
  row.demoFunctionalTestPassed === false));
const blockers = observation.data.blockers.join("\n");
for (const token of [
  "NO_VERIFIED_REVIEWER_ISOLATION",
  "ENABLED",
  "DELETION_CANARY_DISABLED",
  "NO_VERIFIED_FULL_FEATURE_SANDBOX",
  "NO_CONFIRMED_GENERAL_ACCOUNT_DELETION",
  "NO_APPLE_SIGNIN_DELETION_DEVICE_TEST"
]) {
  assert.ok(blockers.includes(token), "Missing release blocker: " + token);
}
assert.equal(observation.data.evidence.ascReviewerAuthObservedOn20261120, "enabled");
assert.equal(observation.data.evidence.ascReviewerFirestoreProfileObservedOn20261120,
  "active_without_reviewer_markers");
assert.equal(observation.data.evidence.separateDeletionCanaryObservedOn20261120,
  "disabled");
assert.ok(!blockers.includes("NO_ACTIVE_VERIFIED_REVIEW_LOGIN"),
  "Do not conflate active ASC login with separate disabled deletion canary");

const release = readReport(["--require-ready"]);
assert.equal(release.code, 1, "Release path must be denied");
assert.equal(release.data.status, "BLOCKED");

// Simulate a lost production feature file without mutating anything in repo.
// The audit must report a structured blocker and continue returning JSON,
// not crash in an assert before the upload guard understands the issue.
const source = fs.readFileSync(program, "utf8");
const captured = [];
const fakeFs = Object.assign({}, fs, {
  existsSync(file) {
    if (path.basename(file) === "management-v2.js") return false;
    return fs.existsSync(file);
  }
});
vm.runInNewContext(source, {
  __dirname: path.dirname(program),
  require(name) { return name === "node:fs" ? fakeFs : require(name); },
  console: {log(line) { captured.push(line); }},
  process: { argv: [process.execPath, program] }
}, { filename: program, timeout: 10000 });
assert.equal(captured.length, 1);
const missing = JSON.parse(captured[0]);
assert.equal(missing.status, "BLOCKED");
assert.ok(missing.blockers.some(b =>
  b.includes("MISSING_FEATURE_SOURCE") && b.includes("management-v2.js")));
assert.equal(missing.coverage.find(c => c.file === "management-v2.js").sourcePresent, false);

console.log("PASS: App Review readiness remains BLOCKED until proven; separate reviewer and deletion-canary states correctly distinguished and missing source reported without crash.");

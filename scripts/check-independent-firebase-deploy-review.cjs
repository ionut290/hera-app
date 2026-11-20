"use strict";
// Offline verification of fail-closed independent approval requirements.
// No GitHub API authentication, production secrets, workflows or deploy calls.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const cp = require("node:child_process");

const expression = '.protection_rules | any(.type == "required_reviewers" and .prevent_self_review == true and ([.reviewers[]?] | length > 0))';
const files = [
  ".github/workflows/deploy-account-deletion-protected.yml",
  ".github/workflows/deploy-account-deletion-request-protected.yml",
  ".github/workflows/prepare-app-review-demo-claims-protected.yml"
];
for (const file of files) {
  const text = fs.readFileSync(file, "utf8");
  assert.ok(text.includes("environment: firebase-account-deletion-production"),
    file + " must use protected environment");
  assert.ok(text.includes(expression),
    file + " must require required_reviewers and prevent_self_review=true");
  assert.match(text, /needs: guard/, file + " must enforce a pre-approval guard");
  assert.match(text, /persist-credentials: false/,
    file + " must never persist checkout credentials");
  assert.doesNotMatch(text, /prevent_self_review\s*==\s*false/,
    file + " must never accept self-approval");
}
const allowed = rule => ({ protection_rules: [rule] });
const required = (reviewers, prevent_self_review) => ({
  type: "required_reviewers", reviewers, prevent_self_review
});
const fakeReviewer = { type: "User", reviewer: { id: 42 } };
const samples = [
  ["independent reviewer", allowed(required([fakeReviewer], true)), true],
  ["empty reviewer list", allowed(required([], true)), false],
  ["self review explicitly allowed", allowed(required([fakeReviewer], false)), false],
  ["missing prevent_self_review", allowed({type:"required_reviewers",reviewers:[fakeReviewer]}), false],
  ["no protection rules", {protection_rules:[]}, false],
  ["no reviews field", {protection_rules:[{type:"branch_policy",prevent_self_review:true}]}, false],
  ["wrong type", allowed({type:"wait_timer",prevent_self_review:true,reviewers:[fakeReviewer]}), false],
  ["missing rules", {}, false]
];
for (const [name, input, permitted] of samples) {
  const result = cp.spawnSync("jq", ["-e", expression], {
    input: JSON.stringify(input), encoding: "utf8"
  });
  assert.ifError(result.error);
  assert.equal(result.status === 0, permitted,
    name + " returned " + result.status + ": " + result.stderr);
}
console.log("PASS: all three sensitive Firebase workflows require an independent reviewer; real jq guard rejects self-review, missing reviewer and missing protection.");

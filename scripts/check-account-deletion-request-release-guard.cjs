"use strict";
// Offline contract for a FUTURE guarded request/status deployment; no credentials or Firebase calls.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const source = fs.readFileSync(".github/workflows/deploy-account-deletion-request-protected.yml","utf8");
const main = fs.readFileSync("functions/main.js","utf8");
const regular = fs.readFileSync(".github/workflows/deploy-firebase-functions.yml","utf8");
assert.match(source, /workflow_dispatch:/);
assert.match(source, /default: DO_NOT_DEPLOY/);
assert.match(source, /DEPLOY_REQUEST_INTAKE_ONLY/);
assert.match(source, /github\.ref == 'refs\/heads\/main'/);
assert.match(source, /SHA_APPROVED.*GITHUB_SHA/);
assert.match(source, /PRIVACY_REVIEW_APPROVED/);
assert.match(source, /FULFILLMENT_PROCESS_VERIFIED/);
assert.match(source, /APPROVAL_REFERENCE/);
assert.match(source, /required_reviewers/);
assert.match(source, /\.prevent_self_review == true/,
  "Never deploy deletion request without independent approval");
assert.match(source, /environment: firebase-account-deletion-production/);
assert.match(source, /needs: guard/);
assert.match(source, /ref: \$\{\{ inputs\.reviewed_sha \}\}/);
assert.match(source, /persist-credentials: false/);
assert.match(source, /project_id.*hera-app-6cd2b/);
const deployment = source.match(/--only functions:[^\n]+/g) || [];
assert.equal(deployment.length,1,"Exactly one function deployment allowed");
assert.equal(deployment[0].trim().replace(/\\$/, "").trim(),
  "--only functions:requestOwnVargaAccountDeletion,functions:getOwnVargaAccountDeletionStatus");
assert.doesNotMatch(source,/--only firestore|--only hosting|functions:deleteOwnVargaAccount|APPROVE_ACCOUNT_DELETION_ONLY/);
assert.doesNotMatch(main,/require\(["']\.\/account-deletion-request["']\)/);
assert.match(regular,/!functions\/account-deletion-request\.js/);
assert.doesNotMatch(source,/__VARGA_ACCOUNT_DELETION_REQUEST_ENABLED\s*=\s*true/);
// The transient functions/main.js MUST export only the two reviewed callables.
// Never dynamically Object.assign ALL exports from an evolving module.
assert.match(source, /const approvedNames = \[/);
assert.match(source, /Object\.keys\(staged\)\.sort\(\)\.join\(','\)/);
assert.match(source, /approvedNames\.some\(name => typeof staged\[name\] !== 'function'\)/);
assert.match(source, /exports\.requestOwnVargaAccountDeletion = stagedAccountDeletionRequest\.requestOwnVargaAccountDeletion/);
assert.match(source, /exports\.getOwnVargaAccountDeletionStatus = stagedAccountDeletionRequest\.getOwnVargaAccountDeletionStatus/);
assert.doesNotMatch(source, /Object\.assign\(exports, require\(["']\.\/account-deletion-request["']\)\)/);
const stageSource = fs.readFileSync("functions/account-deletion-request.js", "utf8");
const stagedNames = [...stageSource.matchAll(/^exports\.(\w+)\s*=\s*onCall\(/gm)].map(x => x[1]).sort();
assert.deepEqual(stagedNames, ["getOwnVargaAccountDeletionStatus","requestOwnVargaAccountDeletion"]);

const guard=source.indexOf("\n  guard:");
const deploy=source.indexOf("\n  deploy:");
assert.ok(guard>=0 && deploy>guard);
assert.ok(source.indexOf("Check that routine deployment") === -1 || source.includes("required_reviewers"));
assert.ok(source.indexOf("Run synthetic privacy guards")<deploy);
assert.ok(source.indexOf("Validate scoped Firebase service account")>deploy);
console.log("PASS: request/status deploy requires signed privacy and fulfillment approvals, protected reviewers, exact source SHA; cannot deploy erasure or enable client UI.");

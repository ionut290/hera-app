"use strict";

/**
 * Solo-maintainer ordinary-release policy. No Firebase SDK, secrets or network.
 * Enforced by the manual workflow BEFORE a job can access any deploy secret.
 *
 * This guard does not approve deletion/retention, change reviewer environments,
 * or enable account erasure. Sensitive workflows remain independently protected.
 */
const assert = require("node:assert/strict");
const fs = require("node:fs");

const PROJECT_REPO = "ionut290/hera-app";
const OWNER = "ionut290";
const ACTION = "DEPLOY_ORDINARY_ONLY";
const PREVIEW = "PREVIEW_ONLY";
const CONFIRM = "CONFIRM_ORDINARY_FIREBASE_DEPLOY";
const shaPattern = /^[0-9a-f]{40}$/;

function evaluate(x) {
  const problems = [];
  if (x.repository !== PROJECT_REPO) problems.push("WRONG_REPOSITORY");
  if (x.actor !== OWNER || (x.triggeringActor && x.triggeringActor !== OWNER)) {
    problems.push("NOT_REPOSITORY_OWNER");
  }
  if (x.event !== "workflow_dispatch" || x.ref !== "refs/heads/main") {
    problems.push("NOT_MANUAL_MAIN");
  }
  if (!shaPattern.test(String(x.inputSha || "")) || x.inputSha !== x.commitSha) {
    problems.push("SHA_MISMATCH");
  }
  if (![PREVIEW,ACTION].includes(x.mode)) problems.push("UNKNOWN_MODE");
  if (x.mode === ACTION) {
    if (x.confirmation !== CONFIRM) problems.push("EXPLICIT_CONFIRMATION_MISSING");
    if (x.attempt !== "1") problems.push("RERUN_REQUIRES_NEW_APPROVAL");
  } else if (x.confirmation !== "DO_NOT_DEPLOY") {
    problems.push("PREVIEW_MUST_NOT_AUTHORIZE_DEPLOY");
  }
  return {allowed:problems.length===0,deploy:x.mode===ACTION && problems.length===0,problems};
}

function validateWorkflowSource() {
  const path = ".github/workflows/deploy-firebase-functions.yml";
  const workflow = fs.readFileSync(path,"utf8");
  assert.match(workflow,/^on:\n  workflow_dispatch:\n/m);
  assert.doesNotMatch(workflow,/^  push:/m,
    "Routine Firebase deployments must never run automatically on a push");
  assert.match(workflow,/default: PREVIEW_ONLY/);
  assert.match(workflow,/CONFIRM_ORDINARY_FIREBASE_DEPLOY/);
  assert.match(workflow,/node scripts\/check-solo-owner-release-policy\.cjs --verify/);
  assert.match(workflow,/^  guard:\n/m);
  assert.match(workflow,/^  deploy:\n    needs: guard\n/m);
  assert.match(workflow,/inputs\.mode == 'DEPLOY_ORDINARY_ONLY'/);
  assert.match(workflow,/persist-credentials: false/);
  assert.match(workflow,/functions:approveUserAccess/);
  assert.doesNotMatch(workflow,/functions:deleteOwnVargaAccount/);
  assert.doesNotMatch(workflow,/functions:requestOwnVargaAccountDeletion/);
  assert.doesNotMatch(workflow,/functions:getOwnVargaAccountDeletionStatus/);
  // The ordinary solo release is limited to reviewed functions.
  // Firestore security rules must be handled via a separate privacy review.
  assert.doesNotMatch(workflow,/firebase deploy\s*\n\s*--only firestore:/);
  assert.doesNotMatch(workflow,/--only firestore:rules,firestore:indexes/);
  assert.doesNotMatch(workflow,/:setIamPolicy|method:\s*["']POST["']/,
    "Ordinary workflow must not widen IAM transport");
  const deployPosition = workflow.indexOf("\n  deploy:\n");
  const secretPosition = workflow.indexOf("secrets.FIREBASE_SERVICE_ACCOUNT_HERA_APP_6CD2B");
  assert.ok(deployPosition >= 0 && secretPosition > deployPosition,
    "No Firebase secret should be consumed by preflight");
  assert.match(workflow,/^permissions:\n  contents: read\n/m);
  assert.doesNotMatch(workflow,/^\s+(contents|actions|id-token|deployments): write\s*$/m);
  return true;
}

function selfTest() {
  const sha = "a".repeat(40);
  const base = {repository:PROJECT_REPO,actor:OWNER,triggeringActor:OWNER,
    event:"workflow_dispatch",ref:"refs/heads/main",inputSha:sha,commitSha:sha,
    attempt:"1",mode:PREVIEW,confirmation:"DO_NOT_DEPLOY"};
  assert.deepEqual(evaluate(base),{allowed:true,deploy:false,problems:[]});
  assert.deepEqual(evaluate({...base,mode:ACTION,confirmation:CONFIRM}),
    {allowed:true,deploy:true,problems:[]});
  const cases = [
    {repository:"other/repo"}, {actor:"other"}, {triggeringActor:"other"},
    {event:"push"}, {ref:"refs/heads/develop"},
    {inputSha:"b".repeat(40)}, {inputSha:"short"}, {commitSha:"b".repeat(40)},
    {mode:"UNKNOWN"}, {mode:ACTION}, {mode:ACTION,confirmation:"DO_NOT_DEPLOY"},
    {mode:ACTION,confirmation:CONFIRM,attempt:"2"},
    {mode:PREVIEW,confirmation:CONFIRM}
  ];
  for (const sample of cases) {
    const request = {...base,...sample};
    assert.equal(evaluate(request).allowed,false,
      "Unsafe manual dispatch was accepted: " + JSON.stringify(sample));
  }
  validateWorkflowSource();
  console.log("PASS: single owner can run PREVIEW, ordinary deploy requires fresh explicit confirmation and SHA; push, privilege escalation, rules deploy, IAM mutation and self reruns are blocked.");
}

if (require.main === module) {
  const mode = process.argv[2];
  if (mode === "--self-test") selfTest();
  else if (mode === "--verify") {
    try {
      validateWorkflowSource();
      const verdict = evaluate({
        repository:process.env.GITHUB_REPOSITORY,actor:process.env.GITHUB_ACTOR,
        triggeringActor:process.env.GITHUB_TRIGGERING_ACTOR,
        event:process.env.GITHUB_EVENT_NAME,ref:process.env.GITHUB_REF,
        inputSha:process.env.INPUT_REVIEWED_SHA,
        commitSha:process.env.GITHUB_SHA,attempt:process.env.GITHUB_RUN_ATTEMPT,
        mode:process.env.INPUT_MODE,confirmation:process.env.INPUT_CONFIRMATION
      });
      if (!verdict.allowed) {
        console.error("Solo release refused:",verdict.problems.join(","));
        process.exitCode = 1;
      } else {
        console.log(verdict.deploy
          ? "Exact SHA and solo-owner ordinary release approved (privacy/destructive functions NOT included)."
          : "Read-only solo-owner preview approved: production actions remain disabled.");
      }
    } catch (_) {
      console.error("Solo owner release guard failed closed.");
      process.exitCode=1;
    }
  } else {
    console.error("Select --self-test or --verify");
    process.exitCode=1;
  }
}
module.exports = { evaluate, validateWorkflowSource };

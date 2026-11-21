#!/usr/bin/env node
"use strict";

// Keep legacy Apple reviewer identities blocked server-side while removing
// all reviewer-only login and demo features from the shipped client.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const cp = require("node:child_process");
const read = (file) => fs.readFileSync(file, "utf8");

const index = read("index.html");
assert.doesNotMatch(index, /apple-review-(?:signin|demo)-link/);

const server = read("functions/apple-review-accounts.js");
assert.match(server, /await assertAdmin\(request.auth\);/);
assert.match(server, /getUserByEmail\(email\)/);
assert.match(server, /setCustomUserClaims\(user.uid, \{ apple_review: true \}\)/);
assert.match(server, /throw new HttpsError\("already-exists"/);
const deploy = read(".github/workflows/deploy-firebase-functions.yml");
assert.match(deploy, /functions:createAppleReviewAccount/);
assert.match(server, /invoker: "public"/);
const main = read("functions/main.js");
assert.match(main, /appleReviewAccountFunctions = require\("\.\/apple-review-accounts"\)/);
assert.match(main, /appleReviewAccountFunctions/);
const netlifyGate = read("netlify/functions/_firebase-token.js");
assert.match(netlifyGate, /user\.apple_review === true/);
assert.match(netlifyGate, /Utente Firebase non autorizzato/);
const guard = read("functions/apple-review-guard.js");
assert.match(guard, /auth\?\.token\?\.apple_review === true/);

const rawModules = {
  "functions/index.js":4,
  "functions/central-notifications.js":1,
  "functions/shared-calendar-view.js":1,
  "functions/shared-operational-views.js":1,
  "functions/whazzup-pdf-drive.js":1,
  "functions/varga-session-transfer.js":1,
  "functions/error-center.js":5,
  "functions/error-reporting.js":1,
  "functions/password-recovery-code.js":1,
  "functions/varga-gestionale-sync.js":4,
  "functions/user-access-approval.js":2
};
for (const [path, count] of Object.entries(rawModules)) {
  const src = read(path);
  assert.match(src, /require\("\.\/apple-review-guard"\)/, "Missing shared server guard: " + path);
  const guardCount = (src.match(/assertNotAppleReviewer\((?:context|request)\?\.auth\)/g) || []).length;
  assert.equal(guardCount, count, "Review guard coverage mismatch in "+path);
}
const firestore = read("firestore.rules"), storage = read("storage.rules");
assert.match(firestore, /appleReviewDemo/);
assert.match(firestore, /token\.get\("apple_review", false\)/);
assert.match(storage, /token\.get\("apple_review", false\)/);

for (const file of [
  "functions/apple-review-accounts.js",
  "functions/apple-review-guard.js",
  ...Object.keys(rawModules)
]) {
  cp.execFileSync(process.execPath, ["--check", file], { stdio:"inherit" });
}
console.log("Legacy reviewer identity remains restricted server-side; client-only reviewer access is removed.");

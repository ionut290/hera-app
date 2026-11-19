"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { execFileSync } = require("node:child_process");

const root = path.resolve(__dirname, "..");
function read(relative) {
  return fs.readFileSync(path.join(root, relative), "utf8");
}
function contains(relative, fragment) {
  assert.ok(read(relative).includes(fragment), relative + " missing: " + fragment);
}

execFileSync(process.execPath, ["--check", path.join(root, "account-privacy-controls.js")]);
execFileSync(process.execPath, ["--check", path.join(root, "functions/account-deletion.js")]);

contains("index.html", 'src="account-privacy-controls.js?');
contains("account-privacy-controls.js", '"auth-apple-login-btn"');
contains("account-privacy-controls.js", 'nativePlatform() === "ios"');
contains("account-privacy-controls.js", "signInWithApple({ skipNativeAuth: true })");
contains("account-privacy-controls.js", "new firebase.auth.OAuthProvider");
contains("account-privacy-controls.js", "reauthenticateWithCredential");
contains("account-privacy-controls.js", '"varga-account-delete-menu"');
contains("account-privacy-controls.js", '"varga-account-delete-pending"');
contains("account-privacy-controls.js", 'httpsCallable("deleteOwnVargaAccount")');
contains("functions/main.js", 'require("./account-deletion")');
contains("functions/account-deletion.js", "request.auth");
contains("functions/account-deletion.js", "auth.token?.auth_time");
contains("functions/account-deletion.js", "await admin.auth().deleteUser(uid)");
contains("functions/account-deletion.js", 'visibility === "personal"');
contains("functions/account-deletion.js", 'path.startsWith(prefix)');
contains("functions/account-deletion.js", "Business records shared with a company are not deleted");
contains("scripts/install-ios-apple-signin.rb", 'configuration.build_settings["IPHONEOS_DEPLOYMENT_TARGET"] = "15.0"');
contains(".github/workflows/ios-appstore-candidate.yml", 'default: "10004"');
contains(".github/workflows/ios-appstore-candidate.yml", '10#$IOS_BUILD_NUMBER < 10004');
contains(".github/workflows/ios-private-testflight.yml", '20000 + GITHUB_RUN_NUMBER');

contains("functions/account-deletion.js", 'unlinkSharedAuthor(db, "chatMessages", "senderId", uid');
contains("functions/account-deletion.js", "await unlinkDeletedChatRecipient(db, uid);");
contains("functions/account-deletion.js", 'unlinkSharedAuthor(db, "appNotifications", "createdByUid", uid');


const standardDeploy = read(".github/workflows/deploy-firebase-functions.yml");
assert.ok(!/--only[^\n]*functions:deleteOwnVargaAccount/.test(standardDeploy),
  "Destructive account deletion must never be included in routine Firebase deployment");
// A merge must NOT push restrictive Firestore chat rules before migrating
// old web/iOS/Android clients. Deploy indexes and rules in separate approvals.
assert.ok(!standardDeploy.includes("--only firestore:rules,firestore:indexes"),
  "Routine function deployment must never update Firestore rules/indexes");
const formerAutoDeploy = read(".github/workflows/deploy-self-registration.yml");
assert.ok(!formerAutoDeploy.includes("npx --yes firebase-tools@latest deploy") &&
  !formerAutoDeploy.includes("credentials_json:") &&
  !formerAutoDeploy.includes("FIREBASE_SERVICE_ACCOUNT_HERA_APP_6CD2B"),
  "Legacy registration workflow must be test-only, never a production Firestore deploy");
const stagedChat = read(".github/workflows/deploy-firestore-chat-protected.yml");
for (const marker of ["workflow_dispatch:", "firebase-private-chat-production",
  "required_reviewers", "APPROVE_INDEXES_ONLY", "APPROVE_RULES_ONLY",
  "CLIENTS_MIGRATED", "PRIVACY_REVIEW_APPROVED",
  "--only firestore:indexes", "--only firestore:rules",
  "node scripts/check-chat-indexes-ready.cjs"]) {
  assert.ok(stagedChat.includes(marker), "Missing staged-chat release guard: " + marker);
}
assert.ok(!stagedChat.includes("  push:"),
  "Staged private-chat production workflow cannot auto-trigger on push");

const protectedRelease = read(".github/workflows/deploy-account-deletion-protected.yml");
assert.ok(protectedRelease.includes("workflow_dispatch:"));
assert.ok(!protectedRelease.includes("  push:"));
assert.ok(protectedRelease.includes("environment: firebase-account-deletion-production"));
assert.ok(protectedRelease.includes("required_reviewers"));
assert.ok(protectedRelease.includes("PRIVACY_REVIEW_APPROVED"));
assert.ok(protectedRelease.includes("DEMO_ACCOUNT_READY"));
assert.ok(protectedRelease.includes("--only functions:deleteOwnVargaAccount"));

const cfg = JSON.parse(read("capacitor.config.json"));
assert.equal(cfg.appId, "it.vargacantieri.hera");
assert.equal(cfg.plugins.FirebaseAuthentication.skipNativeAuth, true);
assert.ok(cfg.plugins.FirebaseAuthentication.providers.includes("apple.com"));
console.log("PASS: App Review account deletion and Apple Sign In source checks.");
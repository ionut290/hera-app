#!/usr/bin/env node
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const rules = fs.readFileSync("firestore.rules", "utf8");
const privacy = fs.readFileSync("privacy.html", "utf8");
const storage = fs.readFileSync("storage.rules", "utf8");

assert.match(rules, /match \/appConfig\/adminUsers\s*\{\s*allow read, write:\s*if isAdmin\(\);\s*\}/);
assert.match(rules, /match \/appConfig\/\{configId\}\s*\{[\s\S]*?configId in \["adminUsers", "driveAdminSecret", "driveBridge"\]/);
assert.match(rules, /match \/operatorPositions\/\{userId\}\s*\{\s*allow read, create, update, delete:\s*if false;/);
assert.match(privacy, /ionut29019@gmail\.com/);
assert.match(privacy, /3892352575/);
assert.match(privacy, /Hera/);
assert.match(privacy, /90 giorni/);
assert.match(privacy, /24 mesi/);
assert.match(privacy, /10 anni/);
assert.match(storage, /match \/lavori-occasionali\/\{plantId\}\/\{ownerUserId\}\/\{documentId\}\/\{fileName\}/);
// Reviewer-only Firebase Auth claim must deny every live Storage path.
assert.match(rules, /function isAppleReviewAccount\(\)\s*\{[\s\S]*?token\.get\("apple_review", false\) == true/);
assert.match(rules, /function signedIn\(\)\s*\{[\s\S]*?!isAppleReviewAccount\(\)/);
assert.match(rules, /match \/appleReviewDemo\/\{demoId\}\s*\{\s*allow read: if isAppleReviewAccount\(\);\s*allow create, update, delete: if false;/);
assert.match(storage, /function isAppleReviewAccount\(\)\s*\{[\s\S]*?token\.get\("apple_review", false\) == true/);
assert.match(storage, /allow read: if request.auth != null && !isAppleReviewAccount\(\) && request.auth.uid == ownerUserId;/);
assert.match(storage, /allow read: if request.auth != null && !isAppleReviewAccount\(\) && \(/);
assert.equal((storage.match(/allow write: if request\.auth != null && !isAppleReviewAccount\(\)/g) || []).length, 2,
  "Both document upload paths must block Apple reviewer writes.");
// A private-message recipient token must never be treated like public access.
const chatRules = rules.match(/match \/chatMessages\/\{messageId\}\s*\{([\s\S]*?)\n    \}/);
assert.ok(chatRules, "Missing explicit chatMessages security rules");
assert.match(chatRules[1], /resource\.data\.get\("recipientId", ""\) == request\.auth\.uid/);
assert.match(chatRules[1], /resource\.data\.get\("senderId", ""\) == request\.auth\.uid/);
assert.doesNotMatch(chatRules[1], /allow read, create, update:\s*if signedIn\(\)/);
const app = fs.readFileSync("app.js", "utf8");
const listener = app.slice(app.indexOf("function subscribeChat()"), app.indexOf("async function notifyForNewChatMessages("));
for (const filter of [
  'where("recipientId", "==", "")',
  'where("recipientId", "==", uid)',
  'where("senderId", "==", uid)'
]) assert.ok(listener.includes(filter), "Missing scoped chat stream: " + filter);
assert.ok(!/\.collection\("chatMessages"\)\s*\.orderBy\("createdAt"/.test(listener),
  "Do not restore the global chat subscription");
const indexes = JSON.parse(fs.readFileSync("firestore.indexes.json", "utf8"));
for (const field of ["recipientId", "senderId"]) {
  assert.ok(indexes.indexes.some((idx) => idx.collectionGroup === "chatMessages" &&
    idx.fields?.[0]?.fieldPath === field && idx.fields?.[1]?.fieldPath === "createdAt"),
  "Missing chat composite index for " + field);
}

console.log("Security policy and privacy content regression checks passed.");

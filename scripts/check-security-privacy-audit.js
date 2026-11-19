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
assert.match(chatRules[1], /resource\.data\.recipientId == request\.auth\.uid/);
assert.match(chatRules[1], /resource\.data\.senderId == request\.auth\.uid/);
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

// Personal GPS coordinates and notification emails must not be readable or
// writable through the legacy global signed-in catchall.
const gpsMatch = rules.match(/match \/gpsUpdateRequests\/\{requestId\} \{([\s\S]*?)\n    \}/);
assert.ok(gpsMatch, "GPS security match missing");
assert.match(gpsMatch[1], /operatorId/);
assert.match(gpsMatch[1], /allow update, delete: if isAdmin\(\)/);
const mailMatch = rules.match(/match \/notificationEmailQueue\/\{emailId\} \{([\s\S]*?)\n    \}/);
assert.ok(mailMatch, "Notification email queue security match missing");
assert.match(mailMatch[1], /allow read, create, update, delete: if isAdmin\(\)/);
const catchAll = rules.slice(rules.lastIndexOf('match /{collection}/{docId} {'));
for (const collection of ['"gpsUpdateRequests"', '"notificationEmailQueue"']) {
  assert.ok(!catchAll.includes(collection), "Sensitive collection in broad catchall: " + collection);
}
assert.ok(app.includes("if (!currentUser || !canManageData()) {") &&
  app.includes("gpsUpdateRequests = [];"),
  "Non-admin client must not run global GPS queue listener");

const appNoticeRules = rules.match(/match \/appNotifications\/\{eventId\} \{([\s\S]*?)\n    \}/);
assert.ok(appNoticeRules, "Missing explicit appNotifications security rule");
for (const pattern of [/createdByUid/, /request\.auth\.uid/, /createdAt/, /allow update, delete: if isAdmin\(\)/]) {
  assert.match(appNoticeRules[1], pattern, "Notification event ownership regression");
}
assert.ok(!catchAll.includes('"appNotifications"'),
  "Global event collection must not inherit unrestricted catchall writes");

const deadlineRules = rules.match(/match \/hoursDeadlineAlerts\/\{alertId\} \{([\s\S]*?)\n    \}/);
assert.ok(deadlineRules, "Missing explicit system deadline alert permissions");
assert.match(deadlineRules[1], /allow create, update, delete: if isAdmin\(\)/);
assert.ok(!catchAll.includes('"hoursDeadlineAlerts"'),
  "Deadline alerts must not inherit broad worker writes");

const ackRules = rules.match(/match \/userAlertAcknowledgements\/\{ackId\} \{([\s\S]*?)\n    \}/);
assert.ok(ackRules, "Missing explicit personal alert acknowledgement rules");
assert.match(ackRules[1], /ackId\.matches/);
assert.match(ackRules[1], /request\.auth\.uid/);
assert.match(ackRules[1], /allow list: if isAdmin\(\)/);
assert.match(ackRules[1], /allow update, delete: if isAdmin\(\)/);
assert.ok(!catchAll.includes('"userAlertAcknowledgements"'),
  "Alert acknowledgement collection must not inherit broad writes and reads");

console.log("Security policy and privacy content regression checks passed.");

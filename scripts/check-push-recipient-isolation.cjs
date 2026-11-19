"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const file = path.join(__dirname, "../functions/user-notifications.js");
const src = fs.readFileSync(file, "utf8");
const normalizeBlock = src.slice(src.indexOf("function normalize(value) {"), src.indexOf("function chunk(items, size) {"));
const resolverBlock = src.slice(src.indexOf("function userMatches(user, rawValue) {"), src.indexOf("async function clearInvalidTokens("));
assert.ok(normalizeBlock.includes('function normalize(value)') && resolverBlock.includes('function resolvePrivateChatRecipient('));
const { normalize, resolveUniquePushUser, resolvePrivateChatRecipient, hasUniquePushTokenOwner } =
  vm.runInNewContext(normalizeBlock + "\n" + resolverBlock +
    "\n({normalize,resolveUniquePushUser,resolvePrivateChatRecipient,hasUniquePushTokenOwner})");

const user = (uid, email, names, pushToken = "synthetic-token") => ({
  uid, email, pushToken,
  labels: new Set([uid, email, ...names].map(normalize))
});
const first = user("uid001", "mario.rossi@example.test", ["Mario Rossi", "Mario"]);
const second = user("uid002", "mario.bianchi@example.test", ["Mario Bianchi", "Mario"]);
const third = user("uid003", "lucia@example.test", ["Lucia Verde"], "");
const users = [first, second, third];

assert.equal(hasUniquePushTokenOwner(users, first), true);
assert.equal(hasUniquePushTokenOwner(users, third), false, "No FCM token means no personal push");
const sharedTokenUsers = [...users, user("uid999", "unrelated@example.test", ["Unrelated"], first.pushToken)];
assert.equal(hasUniquePushTokenOwner(sharedTokenUsers, first), false,
  "Reused FCM token across profiles must never receive a private worker notification");
assert.equal(hasUniquePushTokenOwner(sharedTokenUsers, sharedTokenUsers.at(-1)), false);

assert.equal(resolvePrivateChatRecipient(users, "uid001"), first);
assert.equal(resolvePrivateChatRecipient(users, "mario.rossi@example.test"), first);
assert.equal(resolvePrivateChatRecipient(users, "Mario Rossi"), null,
  "Chat recipient names must never trigger a private push");
assert.equal(resolvePrivateChatRecipient(users, "uid000"), null);
assert.equal(resolvePrivateChatRecipient([...users, user("uid001", "new@example.test", [])], "uid001"), null,
  "Duplicate UID must not lead to an ambiguous private push");
assert.equal(resolveUniquePushUser(users, "Mario"), null,
  "Homonymous workers must never receive a name-only personal push");
assert.equal(resolveUniquePushUser(users, "Mario Rossi"), first);
assert.equal(resolveUniquePushUser(users, "lucia@example.test"), third);
assert.equal(resolveUniquePushUser(users, "Verde"), third,
  "Unique legacy surname may still be resolved when there is no ambiguity");
assert.equal(resolveUniquePushUser(users, "Unknown Operator"), null);
assert.equal(resolveUniquePushUser([...users,
  user("uid004", "other@example.test", ["Lucia Verde"], "token")], "Lucia Verde"), null,
  "Even a no-token identity must disambiguate personal push recipients");

const chat = src.slice(src.indexOf("exports.notifyAndroidChatMessage ="), src.indexOf("function extractHoursRows("));
assert.match(chat, /resolvePrivateChatRecipient\(users, recipientId\)/);
assert.match(chat, /hasUniquePushTokenOwner\(users, directRecipient\)/);
assert.doesNotMatch(chat, /userMatches\(user, recipientId\)/);
const hours = src.slice(src.indexOf("exports.notifyAndroidHoursInserted ="), src.indexOf("function extractSquadraOperators("));
assert.match(hours, /resolveUniquePushUser\(users, row.operatorName\)/);
const squad = src.slice(src.indexOf("exports.notifyAndroidSquadraAssignment ="), src.indexOf("function getCalendarEventParticipants("));
assert.match(squad, /resolveUniquePushUser\(users, operator\)/);
assert.ok(src.includes("}); // Keep profiles without push tokens for ambiguity checks."));
console.log("PASS: private chat UID/email exact matching; hours/squad notifications fail closed on ambiguous names.");

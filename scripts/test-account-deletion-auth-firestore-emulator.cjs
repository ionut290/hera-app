"use strict";
// Runs real Admin Auth + Firestore operations ONLY on the local Firebase emulators.
// Never connect this test to the real VARGA CANTIERI cloud resources.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const projectId = "demo-review-isolation";
if (!process.env.FIRESTORE_EMULATOR_HOST || !process.env.FIREBASE_AUTH_EMULATOR_HOST) {
  throw Error("Refuse to run without BOTH Firestore and Auth local emulators.");
}
if (process.env.GCLOUD_PROJECT && process.env.GCLOUD_PROJECT !== projectId) {
  throw Error("Refuse non-demo Firebase project.");
}
const admin = require("firebase-admin");
admin.initializeApp({ projectId });
const db = admin.firestore();
const auth = admin.auth();
const uid = "integration-delete-user-1";
const otherUid = "integration-keep-user-2";
const removedFiles = [];
const adminHarness = {
  firestore: Object.assign(() => db, { FieldValue: admin.firestore.FieldValue }),
  auth: () => auth,
  storage: () => ({
    bucket: (name) => {
      assert.equal(name, "hera-app-6cd2b.firebasestorage.app");
      return { file: (name) => ({ async delete(options) {
        assert.equal(options.ignoreNotFound, true);
        removedFiles.push(name);
      } }) };
    }
  })
};
class MockHttpsError extends Error {
  constructor(code, message) { super(message); this.code = code; }
}
const sandbox = {
  exports: {},
  require: (id) => {
    if (id === "firebase-admin") return adminHarness;
    if (id === "node:crypto") return require("node:crypto");
    if (id === "firebase-functions/v2/https") {
      return { HttpsError: MockHttpsError, onCall: (_opts, handler) => handler };
    }
    throw Error("Unexpected dependency: " + id);
  },
  console: { error() {} },
  Date
};
vm.runInNewContext(
  fs.readFileSync(path.resolve(__dirname, "../functions/account-deletion.js"), "utf8"),
  sandbox,
  { filename: "functions/account-deletion.js" }
);
const deleteOwn = sandbox.exports.deleteOwnVargaAccount;
const request = (who, data, age = 0) => ({
  data,
  auth: { uid: who, token: { auth_time: Math.floor(Date.now() / 1000) - age } }
});

(async () => {
  // Real Auth emulator users; intentionally no email, so no real-person linkage.
  await auth.createUser({ uid, email: "remove@example.test" });
  await auth.createUser({ uid: otherUid, email: "keep@example.test" });
  const batch = db.batch();
  const record = (collection, id, data) =>
    batch.set(db.collection(collection).doc(id), data);
  record("platformUsers", uid, { state: "test-only" });
  record("platformUsers", otherUid, { state: "must-retain" });
  record("privateDocuments", uid, { note: "synthetic personal" });
  record("documents", "synthetic-file-private", {
    ownerUserId: uid, visibility: "personal",
    storagePath: "documents/" + uid + "/private/file.txt"
  });
  record("documents", "synthetic-file-other", {
    ownerUserId: otherUid, visibility: "personal",
    storagePath: "documents/" + otherUid + "/private/file.txt"
  });
  record("personale", "synthetic-staff", {
    linkedUserId: uid, linkedAuthProvider: "apple.com",
    contractCode: "keep-synthetic-employment-record"
  });
  record("personale", "synthetic-other-staff", {
    linkedUserId: otherUid, contractCode: "keep-other-record"
  });
  record("notifications", "synthetic-shared", {
    actorId: uid, actorName: "Synthetic",
    recipientUserIds: [uid, otherUid],
    readBy: { [uid]: true, [otherUid]: true },
    message: "Retain business info"
  });
  record("chatMessages", "synthetic-chat-owned", {
    senderId: uid, senderEmail: "synthetic@example.test", senderName: "Synthetic Worker",
    recipientId: otherUid, text: "Retain business chat", metadata: { commessaId: "synthetic-job" }
  });
  record("chatMessages", "synthetic-chat-other", {
    senderId: otherUid, senderEmail: "other@example.test", senderName: "Other Worker",
    recipientId: uid, text: "Preserve another worker's message"
  });
  record("appNotifications", "synthetic-app-notice-owned", {
    createdByUid: uid, createdByEmail: "synthetic@example.test", createdByName: "Synthetic Worker",
    title: "Job notice", body: "Retain business notification"
  });
  record("appNotifications", "synthetic-app-notice-other", {
    createdByUid: otherUid, createdByEmail: "other@example.test", title: "Other notice"
  });
  record("notifications", "synthetic-false-state-notice", {
    actorId: otherUid, actorName: "Other Worker",
    recipientUserIds: [otherUid],
    readBy: { [uid]: false, [otherUid]: true },
    deletedBy: { [uid]: false },
    message: "Retained false-state notice"
  });
  record("notificationEmailQueue", "synthetic-email-to-deleted", {
    to: "remove@example.test", createdBy: "admin@test.invalid", status: "pending",
    html: "<p>private synthetic</p>", text: "private synthetic", notificationId: "demo-1"
  });
  record("notificationEmailQueue", "synthetic-email-from-deleted", {
    to: "keep@example.test", createdBy: "remove@example.test", status: "sent",
    text: "Other user's business notice"
  });
  record("notificationEmailQueue", "synthetic-email-other", {
    to: "keep@example.test", createdBy: "other", status: "pending",
    text: "Keep pending for other account"
  });
  record("activityLogs", "synthetic-log-deleted", {
    userId: uid, userEmail: "remove@example.test", userName: "Demo user",
    deviceInfo: "Synthetic test device", actionType: "apertura_app",
    detail: "Keep business action", createdAt: 1
  });
  record("activityLogs", "synthetic-log-other", {
    userId: otherUid, userEmail: "keep@example.test", userName: "Demo other",
    deviceInfo: "Synthetic other device", actionType: "apertura_app",
    createdAt: 2
  });
  record("activityLogs", "synthetic-legacy-email", {
    userEmail: "remove@example.test", userName: "Demo legacy",
    deviceInfo: "Synthetic legacy device", actionType: "login_app",
    createdAt: 3
  });
  record("passwordRecoveryChallenges", "synthetic-uid-challenge", { uid });
  record("passwordRecoveryChallenges", "synthetic-other-challenge", { uid: otherUid });
  await batch.commit();

  await assert.rejects(deleteOwn(request(uid, { confirmation: "INVALID" })),
    (err) => err.code === "invalid-argument");
  await assert.rejects(deleteOwn(request(uid, { confirmation: "ELIMINA" }, 3600)),
    (err) => err.code === "failed-precondition");
  const preflight = await deleteOwn(request(uid, { confirmation: "ELIMINA", mode: "preflight" }));
  assert.equal(preflight.ready, true);
  assert.equal(preflight.personalDocuments, 1);
  assert.equal((await auth.getUser(uid)).uid, uid);
  assert.equal((await db.collection("documents").doc("synthetic-file-private").get()).exists, true);
  const deleted = await deleteOwn(request(uid, { confirmation: "ELIMINA", uid: otherUid }));
  assert.equal(deleted.deleted, true);

  // Firebase Auth Emulator actually deleted only the authenticated synthetic UID.
  await assert.rejects(auth.getUser(uid), (err) => err.code === "auth/user-not-found");
  assert.equal((await auth.getUser(otherUid)).uid, otherUid);
  assert.equal((await db.doc("platformUsers/" + uid).get()).exists, false);
  assert.equal((await db.doc("platformUsers/" + otherUid).get()).exists, true);
  assert.equal((await db.doc("privateDocuments/" + uid).get()).exists, false);
  assert.equal((await db.doc("documents/synthetic-file-private").get()).exists, false);
  assert.equal((await db.doc("documents/synthetic-file-other").get()).exists, true);
  assert.equal((await db.doc("passwordRecoveryChallenges/synthetic-uid-challenge").get()).exists, false);
  assert.equal((await db.doc("passwordRecoveryChallenges/synthetic-other-challenge").get()).exists, true);
  const staff = (await db.doc("personale/synthetic-staff").get()).data();
  assert.equal(staff.linkedUserId, undefined);
  assert.equal(staff.contractCode, "keep-synthetic-employment-record");
  const otherStaff = (await db.doc("personale/synthetic-other-staff").get()).data();
  assert.equal(otherStaff.linkedUserId, otherUid);
  const notification = (await db.doc("notifications/synthetic-shared").get()).data();
  assert.equal(notification.actorId, undefined);
  assert.equal(notification.actorName, "Account eliminato");
  assert.deepEqual(notification.recipientUserIds, [otherUid]);
  assert.deepEqual(notification.readBy, { [otherUid]: true });
  assert.equal(notification.message, "Retain business info");
  const falseStateNotice = (await db.doc("notifications/synthetic-false-state-notice").get()).data();
  assert.equal(falseStateNotice.actorId, otherUid);
  assert.deepEqual(falseStateNotice.recipientUserIds, [otherUid]);
  assert.deepEqual(falseStateNotice.readBy, { [otherUid]: true });
  assert.deepEqual(falseStateNotice.deletedBy, {});
  assert.equal(falseStateNotice.message, "Retained false-state notice");
  const companyChat = (await db.doc("chatMessages/synthetic-chat-owned").get()).data();
  assert.equal(companyChat.senderId, undefined);
  assert.equal(companyChat.senderEmail, undefined);
  assert.equal(companyChat.senderName, "Account eliminato");
  assert.equal(companyChat.recipientId, otherUid);
  assert.equal(companyChat.text, "Retain business chat");
  assert.deepEqual(companyChat.metadata, { commessaId: "synthetic-job" });
  const otherChat = (await db.doc("chatMessages/synthetic-chat-other").get()).data();
  assert.equal(otherChat.senderId, otherUid);
  assert.equal(otherChat.senderEmail, "other@example.test");
  assert.equal(otherChat.text, "Preserve another worker's message");
  assert.match(otherChat.recipientId, /^deleted-recipient-[0-9a-f-]{36}$/);
  assert.notEqual(otherChat.recipientId, uid);
  const ownedNotice = (await db.doc("appNotifications/synthetic-app-notice-owned").get()).data();
  assert.equal(ownedNotice.createdByUid, undefined);
  assert.equal(ownedNotice.createdByEmail, undefined);
  assert.equal(ownedNotice.createdByName, "Account eliminato");
  assert.equal(ownedNotice.body, "Retain business notification");
  const untouchedNotice = (await db.doc("appNotifications/synthetic-app-notice-other").get()).data();
  assert.equal(untouchedNotice.createdByUid, otherUid);
  assert.equal(untouchedNotice.createdByEmail, "other@example.test");
  const emailRemoved = (await db.doc("notificationEmailQueue/synthetic-email-to-deleted").get()).data();
  assert.equal(emailRemoved.to, undefined);
  assert.equal(emailRemoved.html, undefined);
  assert.equal(emailRemoved.text, undefined);
  assert.equal(emailRemoved.status, "cancelled_account_deleted");
  const emailFromRemoved = (await db.doc("notificationEmailQueue/synthetic-email-from-deleted").get()).data();
  assert.equal(emailFromRemoved.to, "keep@example.test");
  assert.equal(emailFromRemoved.createdBy, "Account eliminato");
  assert.equal(emailFromRemoved.text, "Other user's business notice");
  const otherEmail = (await db.doc("notificationEmailQueue/synthetic-email-other").get()).data();
  assert.equal(otherEmail.to, "keep@example.test");
  assert.equal(otherEmail.status, "pending");
  assert.equal(otherEmail.createdBy, "other");

  const logRemoved = (await db.doc("activityLogs/synthetic-log-deleted").get()).data();
  assert.equal(logRemoved.userId, undefined);
  assert.equal(logRemoved.userEmail, undefined);
  assert.equal(logRemoved.userName, "Account eliminato");
  assert.equal(logRemoved.deviceInfo, undefined);
  assert.equal(logRemoved.detail, "Keep business action");
  const otherLog = (await db.doc("activityLogs/synthetic-log-other").get()).data();
  assert.equal(otherLog.userId, otherUid);
  assert.equal(otherLog.userEmail, "keep@example.test");
  assert.equal(otherLog.deviceInfo, "Synthetic other device");
  const legacyLog = (await db.doc("activityLogs/synthetic-legacy-email").get()).data();
  assert.equal(legacyLog.userEmail, undefined);
  assert.equal(legacyLog.userName, "Account eliminato");
  assert.equal(legacyLog.deviceInfo, undefined);
  assert.deepEqual(removedFiles, ["documents/" + uid + "/private/file.txt"]);

  console.log("PASS: REAL Auth+Firestore EMULATORS: one test account deleted; other login, files, personnel, notifications and business history protected.");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
}).finally(async () => { await admin.app().delete(); });

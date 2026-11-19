"use strict";
// Runs ONLY against the Firebase Firestore emulator and synthetic demo identities.
// This script must never be used against the production project.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

if (!process.env.FIRESTORE_EMULATOR_HOST) {
  throw new Error("Refusing to run account deletion integration test outside Firestore Emulator");
}
if (process.env.GCLOUD_PROJECT && process.env.GCLOUD_PROJECT !== "demo-review-isolation") {
  throw new Error("Refusing to use a non-demo Firebase project");
}
const admin = require("firebase-admin");
admin.initializeApp({ projectId: "demo-review-isolation" });
const db = admin.firestore();
const userId = "synthetic-delete-worker-1";
const otherId = "synthetic-preserved-worker-2";
const deletedAuth = [];
const removedStorage = [];
let throwOnGetUser = false;

class MockHttpsError extends Error {
  constructor(code, message) { super(message); this.code = code; }
}
const adminHarness = {
  firestore: Object.assign(() => db, { FieldValue: admin.firestore.FieldValue }),
  auth: () => ({
    getUser: async (uid) => {
      if (throwOnGetUser) throw new Error("Test Auth backend unavailable");
      if (uid !== userId && uid !== otherId) throw new Error("Unknown synthetic user");
      return { uid };
    },
    deleteUser: async (uid) => { deletedAuth.push(uid); }
  }),
  storage: () => ({
    bucket: (name) => {
      assert.equal(name, "hera-app-6cd2b.firebasestorage.app");
      return { file: (name) => ({
        async delete({ ignoreNotFound }) {
          assert.equal(ignoreNotFound, true);
          removedStorage.push(name);
        }
      }) };
    }
  })
};
const sandbox = {
  exports: {},
  require: (name) => {
    if (name === "firebase-admin") return adminHarness;
    if (name === "firebase-functions/v2/https") return {
      HttpsError: MockHttpsError,
      onCall: (_opts, handler) => handler
    };
    throw Error("Unexpected dependency: " + name);
  },
  Date,
  console: { error() {} }
};
const code = fs.readFileSync(path.resolve(__dirname, "../functions/account-deletion.js"), "utf8");
vm.runInNewContext(code, sandbox, { filename: "functions/account-deletion.js" });
const handler = sandbox.exports.deleteOwnVargaAccount;

async function exists(collection, id) {
  return (await db.collection(collection).doc(id).get()).exists;
}
async function seed() {
  const batch = db.batch();
  function add(collection, id, data) { batch.set(db.collection(collection).doc(id), data); }
  add("documents", "synthetic-private", {
    ownerUserId: userId, visibility: "personal",
    storagePath: "documents/" + userId + "/p/private.pdf",
    thumbnailPath: "documents/" + userId + "/p/thumbnail.jpg"
  });
  add("documents", "synthetic-legacy", {
    ownerUserId: userId,
    storagePath: "documents/" + userId + "/old/legacy.pdf"
  });
  add("documents", "synthetic-shared", {
    ownerUserId: userId, visibility: "shared",
    storagePath: "documents/" + userId + "/shared/shared.pdf",
    thumbnailPath: "documents/" + userId + "/p/thumbnail.jpg"
  });
  add("documents", "synthetic-other", {
    ownerUserId: otherId, visibility: "personal",
    storagePath: "documents/" + otherId + "/p/other.pdf"
  });
  add("platformUsers", userId, { name: "TEST IDENTITY, NOT A REAL PERSON" });
  add("platformUsers", otherId, { name: "SYNTHETIC OTHER IDENTITY" });
  add("privateDocuments", userId, { note: "test private data" });
  add("userOperatorLinks", userId, { test: true });
  add("operatorPositions", userId, { test: true });
  add("passwordRecoveryChallenges", "synthetic-challenge-private", { uid: userId });
  add("passwordRecoveryChallenges", "synthetic-challenge-other", { uid: otherId });
  add("personale", "synthetic-employee", {
    linkedUserId: userId, linkedAuthProvider: "google.com",
    googleDisplayName: "TEST", emailVerified: true,
    profilePhotoSource: "google", photoURL: "https://example.test/google.jpg",
    businessContract: "SYNTHETIC KEEP"
  });
  add("personale", "synthetic-other-employee", {
    linkedUserId: otherId, businessContract: "DO NOT TOUCH"
  });
  add("oreReports", "synthetic-timesheet", {
    userId, workedHours: 8, marker: "RETAIN BUSINESS RECORD"
  });
  await batch.commit();
}
function req(uid, data) {
  return { auth: { uid, token: { auth_time: Math.floor(Date.now() / 1000) } }, data };
}
async function rejectsCode(promise, code) {
  await assert.rejects(promise, (error) => error?.code === code);
}
(async () => {
  await seed();
  await rejectsCode(handler({ data: { confirmation: "ELIMINA" } }), "unauthenticated");
  await rejectsCode(handler(req(userId, { confirmation: "INVALID" })), "invalid-argument");
  const preflight = await handler(req(userId, { confirmation: "ELIMINA", mode: "preflight" }));
  assert.equal(preflight.ready, true);
  assert.equal(preflight.personalDocuments, 2);
  assert.deepEqual(deletedAuth, []);
  assert.deepEqual(removedStorage, []);
  assert.equal(await exists("documents", "synthetic-private"), true);
  throwOnGetUser = true;
  await rejectsCode(handler(req(userId, { confirmation: "ELIMINA" })), "internal");
  throwOnGetUser = false;
  assert.deepEqual(deletedAuth, []);
  assert.equal(await exists("documents", "synthetic-private"), true);
  const result = await handler(req(userId, { confirmation: "ELIMINA" }));
  assert.equal(result.deleted, true);
  assert.deepEqual(deletedAuth, [userId]);
  assert.equal(await exists("documents", "synthetic-private"), false);
  assert.equal(await exists("documents", "synthetic-legacy"), false);
  assert.equal(await exists("documents", "synthetic-shared"), true);
  assert.equal(await exists("documents", "synthetic-other"), true);
  assert.equal(await exists("platformUsers", userId), false);
  assert.equal(await exists("platformUsers", otherId), true);
  assert.equal(await exists("privateDocuments", userId), false);
  assert.equal(await exists("passwordRecoveryChallenges", "synthetic-challenge-private"), false);
  assert.equal(await exists("passwordRecoveryChallenges", "synthetic-challenge-other"), true);
  assert.equal(await exists("oreReports", "synthetic-timesheet"), true);
  const person = (await db.collection("personale").doc("synthetic-employee").get()).data();
  assert.equal(person.linkedUserId, undefined);
  assert.equal(person.googleDisplayName, undefined);
  assert.equal(person.photoURL, undefined);
  assert.equal(person.businessContract, "SYNTHETIC KEEP");
  const other = (await db.collection("personale").doc("synthetic-other-employee").get()).data();
  assert.equal(other.linkedUserId, otherId);
  assert.equal(other.businessContract, "DO NOT TOUCH");
  assert.deepEqual(removedStorage.sort(), [
    "documents/" + userId + "/p/private.pdf",
    "documents/" + userId + "/old/legacy.pdf"
  ].sort());
  // The private thumbnail remains because an explicitly retained shared record uses it.
  assert.equal(removedStorage.includes("documents/" + userId + "/p/thumbnail.jpg"), false);
  // Fail closed if historical documents have no safely identifiable Storage path.
  await db.collection("documents").doc("synthetic-unmapped").set({
    ownerUserId: otherId, visibility: "personal", fileUrl: "https://example.test/unmanaged"
  });
  await rejectsCode(handler(req(otherId, { confirmation: "ELIMINA", mode: "preflight" })),
    "failed-precondition");
  assert.deepEqual(deletedAuth, [userId]);
  assert.equal(await exists("platformUsers", otherId), true);
  console.log("PASS: Real Firestore emulator: auth isolation, preflight, legacy cleanup, retained shared files, personnel unlink, business preservation.");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
}).finally(async () => { await admin.app().delete(); });

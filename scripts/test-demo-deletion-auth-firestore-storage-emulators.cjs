"use strict";
/*
 * Execute the ACTUAL deployed-candidate deletion handler against local
 * Firebase Auth, Firestore AND Cloud Storage emulators. No mocked deletions.
 * No credentials, production services, real employee data or live deletion.
 */
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const projectId = "demo-deletion-canary-storage";
for (const name of ["FIREBASE_AUTH_EMULATOR_HOST", "FIRESTORE_EMULATOR_HOST", "FIREBASE_STORAGE_EMULATOR_HOST"]) {
  const value = process.env[name] || "";
  if (!/^(127\.0\.0\.1|localhost|\[::1\]):[0-9]{2,5}$/.test(value)) {
    throw Error("All Firebase emulators must use loopback: " + name);
  }
}
for (const name of ["GCLOUD_PROJECT", "GOOGLE_CLOUD_PROJECT", "FIREBASE_PROJECT"]) {
  if (process.env[name] && process.env[name] !== projectId) {
    throw Error("Production project forbidden: " + name);
  }
}
if (process.env.GOOGLE_APPLICATION_CREDENTIALS) {
  throw Error("Service account credentials forbidden");
}
const admin = require("../functions/node_modules/firebase-admin");
admin.initializeApp({ projectId });
const db = admin.firestore();
const auth = admin.auth();
const now = Math.floor(Date.now() / 1000);
const demoUid = "demo-canary-" + process.pid;
const otherUid = "other-canary-" + process.pid;
// This real Admin Storage client uses the emulator URL from the enforced
// FIREBASE_STORAGE_EMULATOR_HOST variable. The bucket is a local-emulator
// namespace only. No service account credentials or remote endpoints exist.
const bucket = admin.storage().bucket("hera-app-6cd2b.firebasestorage.app");
const ownPath = "documents/" + demoUid + "/demo/private.pdf";
const otherPath = "documents/" + otherUid + "/other/keep.pdf";
const retainedPath = "documents/" + demoUid + "/demo/retained.pdf";
const adminHarness = {
  firestore: Object.assign(() => db, { FieldValue: admin.firestore.FieldValue }),
  auth: () => auth,
  storage: () => admin.storage()
};
class HttpsError extends Error {
  constructor(code, message) { super(message); this.code = code; }
}
const mod = { exports: {} };
const factory = vm.runInThisContext(
  "(function (exports, require, console) {\n" +
    fs.readFileSync("functions/account-deletion.js", "utf8") + "\n})",
  { filename: "functions/account-deletion.js" }
);
factory(mod.exports, name => {
  if (name === "firebase-admin") return adminHarness;
  if (name === "node:crypto") return require("node:crypto");
  if (name === "firebase-functions/v2/https") return {
    HttpsError,
    onCall(options, handler) {
      assert.equal(options.region, "europe-west1");
      assert.ok(options.maxInstances <= 3);
      return handler;
    }
  };
  throw Error("Unexpected production dependency: " + name);
}, { error() {} });
const invoke = mod.exports.deleteOwnVargaAccount;
assert.equal(typeof invoke, "function");
function call(uid, data, claims = { appReviewDemo: true, apple_review: true }, age = 0) {
  return { auth: { uid, token: { ...claims, auth_time: now - age } }, data };
}
async function denied(promise, code) {
  await assert.rejects(promise, e => e instanceof HttpsError && e.code === code);
}
async function verifyPreserved() {
  assert.equal((await auth.getUser(demoUid)).uid, demoUid);
  assert.equal((await auth.getUser(otherUid)).uid, otherUid);
  assert.equal((await db.doc("platformUsers/" + demoUid).get()).exists, true);
  assert.equal((await db.doc("platformUsers/" + otherUid).get()).exists, true);
  assert.equal((await db.doc("documents/demo-file").get()).exists, true);
  assert.equal((await db.doc("documents/other-file").get()).exists, true);
  for (const path of [ownPath, otherPath, retainedPath]) {
    assert.equal((await bucket.file(path).exists())[0], true,
      "Must preserve emulator object before authorized deletion: " + path);
  }
}
async function main() {
  const users = [demoUid, otherUid];
  try {
    // The whitelisted address exists ONLY inside the disposable Auth Emulator.
    await auth.createUser({ uid: demoUid, email: "sqadrahera+varga-demo@gmail.com" });
    await auth.createUser({ uid: otherUid, email: "other-user@example.test" });
    await auth.setCustomUserClaims(demoUid, {
      appReviewDemo: true, apple_review: true
    });
    const batch = db.batch();
    batch.set(db.doc("platformUsers/" + demoUid), { fixture: "demo" });
    batch.set(db.doc("platformUsers/" + otherUid), { fixture: "keep" });
    batch.set(db.doc("privateDocuments/" + demoUid), { fixture: "owned" });
    batch.set(db.doc("documents/demo-file"), {
      ownerUserId: demoUid, visibility: "personal",
      storagePath: ownPath
    });
    batch.set(db.doc("documents/other-file"), {
      ownerUserId: otherUid, visibility: "personal",
      storagePath: otherPath
    });
    // A personal object still referenced by a company document is retained.
    batch.set(db.doc("documents/demo-retained-personal"), {
      ownerUserId: demoUid, visibility: "personal", storagePath: retainedPath
    });
    batch.set(db.doc("documents/company-shared-attachment"), {
      ownerUserId: otherUid, visibility: "shared", storagePath: retainedPath
    });
    batch.set(db.doc("chatMessages/demo-message"), {
      senderId: demoUid, senderEmail: "sqadrahera+varga-demo@gmail.com",
      senderName: "Demo", recipientId: otherUid,
      text: "Preserve operational text"
    });
    batch.set(db.doc("chatMessages/other-message"), {
      senderId: otherUid, senderEmail: "other-user@example.test",
      recipientId: demoUid, text: "Preserve other user message"
    });
    batch.set(db.doc("oreReports/company-record"), {
      userId: demoUid, hours: 8, fixture: "must remain"
    });
    await batch.commit();
    // Actual file bytes exist ONLY inside Cloud Storage emulator, not GCS.
    for (const [objectPath, body] of [
      [ownPath, "private synthetic attachment"],
      [otherPath, "other user's preserved attachment"],
      [retainedPath, "company-linked shared attachment"]
    ]) {
      await bucket.file(objectPath).save(Buffer.from(body), {
        resumable: false,
        contentType: "application/pdf"
      });
    }
    await denied(invoke(call(demoUid, { confirmation: "ELIMINA" }, {})), "permission-denied");
    await denied(invoke(call(otherUid, { confirmation: "ELIMINA" })), "permission-denied");
    await denied(invoke(call(demoUid, { confirmation: "ELIMINA" }, undefined, 601)), "failed-precondition");
    await denied(invoke(call(demoUid, { confirmation: "ELIMINA", uid: otherUid })), "invalid-argument");
    await denied(invoke(call(demoUid, { confirmation: "ELIMINA", mode: "unexpected" })), "invalid-argument");
    await verifyPreserved();

    await auth.updateUser(demoUid, { disabled: true });
    await denied(invoke(call(demoUid, { confirmation: "ELIMINA", mode: "preflight" })), "permission-denied");
    await denied(invoke(call(demoUid, { confirmation: "ELIMINA" })), "permission-denied");
    await verifyPreserved();
    await auth.updateUser(demoUid, { disabled: false });

    await auth.setCustomUserClaims(demoUid, { appReviewDemo: true });
    await denied(invoke(call(demoUid, { confirmation: "ELIMINA" })), "permission-denied");
    await verifyPreserved();
    await auth.setCustomUserClaims(demoUid, {
      appReviewDemo: true, apple_review: true
    });

    const preview = await invoke(call(demoUid, {
      confirmation: "ELIMINA", mode: "preflight"
    }));
    assert.equal(preview.ready, true);
    assert.equal(preview.personalDocuments, 2);
    await verifyPreserved();

    // ONLY local emulator fixture is now irreversibly deleted.
    const result = await invoke(call(demoUid, { confirmation: "ELIMINA" }));
    assert.equal(result.deleted, true);
    await assert.rejects(auth.getUser(demoUid),
      error => error.code === "auth/user-not-found");
    assert.equal((await auth.getUser(otherUid)).uid, otherUid);
    assert.equal((await db.doc("platformUsers/" + demoUid).get()).exists, false);
    assert.equal((await db.doc("platformUsers/" + otherUid).get()).exists, true);
    assert.equal((await db.doc("privateDocuments/" + demoUid).get()).exists, false);
    assert.equal((await db.doc("documents/demo-file").get()).exists, false);
    assert.equal((await db.doc("documents/other-file").get()).exists, true);
    assert.equal((await db.doc("documents/demo-retained-personal").get()).exists, false);
    assert.equal((await db.doc("documents/company-shared-attachment").get()).exists, true);
    assert.equal((await db.doc("oreReports/company-record").get()).data().fixture, "must remain");
    assert.equal((await bucket.file(ownPath).exists())[0], false, "Private file must be deleted");
    assert.equal((await bucket.file(otherPath).exists())[0], true, "Other user's file must survive");
    assert.equal((await bucket.file(retainedPath).exists())[0], true,
      "A file referenced by retained company metadata must survive");
    const retainedBytes = (await bucket.file(retainedPath).download())[0].toString();
    assert.equal(retainedBytes, "company-linked shared attachment");
    const chat = (await db.doc("chatMessages/demo-message").get()).data();
    assert.equal(chat.senderId, undefined);
    assert.equal(chat.senderEmail, undefined);
    assert.equal(chat.text, "Preserve operational text");
    const otherChat = (await db.doc("chatMessages/other-message").get()).data();
    assert.equal(otherChat.senderId, otherUid);
    assert.match(otherChat.recipientId, /^deleted-recipient-[0-9a-f-]{36}$/);
    assert.equal(otherChat.text, "Preserve other user message");
    console.log("PASS: local Auth+Firestore+Storage emulators; actual private Storage object removed, retained business and other-user objects preserved");
  } finally {
    for (const uid of users) {
      try { await auth.deleteUser(uid); }
      catch (error) { if (error.code !== "auth/user-not-found") throw error; }
    }
    await admin.app().delete();
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });

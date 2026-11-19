"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

class MockHttpsError extends Error {
  constructor(code, message) {
    super(message);
    this.code = code;
  }
}

function setup(options = {}) {
  const injected = { storage: false, firestore: false, personnel: false, auth: false };
  const users = new Set(["worker-1", "worker-2"]);
  const docs = [
    { collection: "documents", id: "private-1", ownerUserId: "worker-1", visibility: "personal", storagePath: "documents/worker-1/doc1/private.pdf", thumbnailPath: "documents/worker-1/doc1/thumbnail.jpg" },
    { collection: "documents", id: "shared-1", ownerUserId: "worker-1", visibility: "shared", storagePath: "documents/worker-1/doc2/shared.pdf" },
    { collection: "documents", id: "other-1", ownerUserId: "worker-2", visibility: "personal", storagePath: "documents/worker-2/doc3/other.pdf" },
    { collection: "documentReminderAcknowledgements", id: "ack-1", userId: "worker-1" },
    { collection: "documentReminderAcknowledgements", id: "ack-2", userId: "worker-2" },
    { collection: "passwordRecoveryChallenges", id: "challenge-1", uid: "worker-1" },
    { collection: "passwordRecoveryChallenges", id: "challenge-2", uid: "worker-2" },
    { collection: "platformUsers", id: "worker-1" },
    { collection: "platformUsers", id: "worker-2" },
    { collection: "privateDocuments", id: "worker-1" },
    { collection: "notifications", id: "shared-message", actorId: "worker-1",
      actorName: "Deleted Worker", recipientUserIds: ["worker-1", "worker-2"],
      readBy: { "worker-1": true, "worker-2": true },
      deletedBy: { "worker-1": true }, message: "Retained business notification" },
    { collection: "notifications", id: "other-message", actorId: "worker-2",
      actorName: "Other Worker", recipientUserIds: ["worker-2"],
      readBy: { "worker-2": true }, message: "Another user's business notification" },
    { collection: "chatMessages", id: "chat-from-deleted", senderId: "worker-1",
      senderEmail: "synthetic@example.test", senderName: "Synthetic Worker",
      recipientId: "worker-2", text: "Company chat should remain", metadata: { commessaId: "job-test" } },
    { collection: "chatMessages", id: "chat-from-other", senderId: "worker-2",
      senderEmail: "other@example.test", senderName: "Other Worker",
      recipientId: "worker-1", text: "Another person's message" },
    { collection: "appNotifications", id: "app-created-by-deleted", createdByUid: "worker-1",
      createdByEmail: "synthetic@example.test", createdByName: "Synthetic Worker",
      title: "Synthetic job notice", body: "Keep work information" },
    { collection: "appNotifications", id: "app-created-by-other", createdByUid: "worker-2",
      createdByEmail: "other@example.test", title: "Another person's alert" },
    { collection: "personale", id: "employee-1", linkedUserId: "worker-1", linkedAuthProvider: "google.com", googleLastLoginAt: 123, googleDisplayName: "Test", emailVerified: true, profilePhotoSource: "google", photoURL: "https://example.test/google.jpg", contractCode: "business-retained" },
    { collection: "personale", id: "employee-2", linkedUserId: "worker-2", linkedAuthProvider: "google.com", contractCode: "other-employee" },
    { collection: "personale", id: "employee-3", linkedUserId: "worker-1", linkedAuthProvider: "apple.com", profilePhotoSource: "operator", photoURL: "https://example.test/business.jpg", contractCode: "another-retained" }
  ];
  const removed = [];
  const removedFiles = [];
  const removedUsers = [];
  const updated = [];
  const db = {
    collection(collection) {
      return {
        where(field, op, value) {
          assert.ok(["==", "array-contains"].includes(op));
          return {
            async get() {
              return {
                docs: docs.filter((d) => {
                  const actual = field.split(".").reduce((obj, key) => obj?.[key], d);
                  return d.collection === collection && (op === "array-contains"
                    ? Array.isArray(actual) && actual.includes(value) : actual === value);
                })
                  .map((d) => ({
                    id: d.id,
                    ref: {
                      path: d.collection + "/" + d.id,
                      async update(fields) {
                        if (options.failPersonnelOnce && !injected.personnel) {
                          injected.personnel = true;
                          throw new Error("Injected personnel write failure");
                        }
                        updated.push({ path: this.path, fields: Object.keys(fields) });
                        for (const [key, value] of Object.entries(fields)) {
                          const parts = key.split(".");
                          const parent = parts.length === 2 ? d[parts[0]] : d;
                          const field = parts.length === 2 ? parts[1] : key;
                          if (value === Symbol.for("firebase-field-delete")) delete parent[field];
                          else if (value?.arrayRemove) {
                            parent[field] = (parent[field] || []).filter((item) => !value.arrayRemove.includes(item));
                          } else parent[field] = value;
                        }
                      }
                    },
                    data: () => d
                  }))
              };
            }
          };
        }
      };
    },
    doc(docPath) {
      return { path: docPath };
    },
    async recursiveDelete(ref) {
      if (options.failFirestoreOnce && !injected.firestore) {
        injected.firestore = true;
        throw new Error("Injected Firestore deletion failure");
      }
      removed.push(ref.path);
    }
  };
  const mockFirestore = () => db;
  mockFirestore.FieldValue = {
    delete: () => Symbol.for("firebase-field-delete"),
    arrayRemove: (...values) => ({ arrayRemove: values })
  };
  const admin = {
    firestore: mockFirestore,
    auth: () => ({
      getUser: async (uid) => {
        if (!users.has(uid)) throw new Error("not found");
        return { uid };
      },
      deleteUser: async (uid) => {
        if (options.failAuthOnce && !injected.auth) {
          injected.auth = true;
          throw new Error("Injected Firebase Auth deletion failure");
        }
        assert.ok(users.has(uid));
        removedUsers.push(uid);
        users.delete(uid);
      }
    }),
    storage: () => ({
      bucket: () => ({
        file: (name) => ({
          async delete(deleteOptions) {
            assert.equal(deleteOptions.ignoreNotFound, true);
            if (options.failStorageOnce && !injected.storage) {
              injected.storage = true;
              throw new Error("Injected Storage deletion failure");
            }
            removedFiles.push(name);
          }
        })
      })
    })
  };
  let handler;
  const sandbox = {
    exports: {},
    require: (id) => {
      if (id === "firebase-admin") return admin;
      if (id === "firebase-functions/v2/https") {
        return { onCall: (_opts, callback) => callback, HttpsError: MockHttpsError };
      }
      throw new Error("Unexpected module: " + id);
    },
    console: { error: () => {} },
    Date
  };
  const code = fs.readFileSync(path.resolve(__dirname, "../functions/account-deletion.js"), "utf8");
  vm.runInNewContext(code, sandbox, { filename: "functions/account-deletion.js" });
  handler = sandbox.exports.deleteOwnVargaAccount;
  assert.equal(typeof handler, "function");
  return { handler, docs, removed, removedFiles, removedUsers, updated };
}

async function rejectedCode(promise, expected) {
  await assert.rejects(promise, (error) => error?.code === expected);
}

(async () => {
  const goodAuth = { uid: "worker-1", token: { auth_time: Math.floor(Date.now() / 1000) } };
  let state = setup();
  await rejectedCode(state.handler({ data: { confirmation: "ELIMINA" } }), "unauthenticated");
  await rejectedCode(state.handler({ auth: goodAuth, data: {} }), "invalid-argument");
  await rejectedCode(state.handler({
    auth: { uid: "worker-1", token: { auth_time: Math.floor(Date.now() / 1000) - 7200 } },
    data: { confirmation: "ELIMINA" }
  }), "failed-precondition");
  assert.equal(state.removedUsers.length, 0);

  state = setup();
  const ready = await state.handler({ auth: goodAuth, data: { confirmation: "ELIMINA", mode: "preflight" } });
  assert.equal(ready.ready, true);
  assert.equal(ready.personalDocuments, 1);
  assert.deepEqual(state.removedUsers, []);
  assert.deepEqual(state.removedFiles, []);
  assert.deepEqual(state.removed, []);
  assert.deepEqual(state.updated, []);

  state = setup();
  const result = await state.handler({
    auth: goodAuth,
    data: { confirmation: "ELIMINA", uid: "worker-2" }
  });
  assert.equal(result.deleted, true);
  assert.deepEqual(state.removedUsers, ["worker-1"]);
  assert.deepEqual(state.removedFiles, ["documents/worker-1/doc1/private.pdf", "documents/worker-1/doc1/thumbnail.jpg"]);
  assert.ok(state.removed.includes("documents/private-1"));
  assert.ok(state.removed.includes("platformUsers/worker-1"));
  assert.ok(state.removed.includes("privateDocuments/worker-1"));
  // Employee registry records survive, but the deleted login is unlinked.
  const employee = state.docs.find((d) => d.id === "employee-1");
  assert.equal(employee.linkedUserId, undefined);
  assert.equal(employee.linkedAuthProvider, undefined);
  assert.equal(employee.photoURL, undefined);
  assert.equal(employee.contractCode, "business-retained");
  const businessPortrait = state.docs.find((d) => d.id === "employee-3");
  assert.equal(businessPortrait.linkedUserId, undefined);
  assert.equal(businessPortrait.photoURL, "https://example.test/business.jpg");
  assert.equal(businessPortrait.contractCode, "another-retained");
  const otherEmployee = state.docs.find((d) => d.id === "employee-2");
  assert.equal(otherEmployee.linkedUserId, "worker-2");
  assert.equal(state.updated.length, 6);
  const priorChat = state.docs.find((d) => d.id === "chat-from-deleted");
  assert.equal(priorChat.senderId, undefined);
  assert.equal(priorChat.senderEmail, undefined);
  assert.equal(priorChat.senderName, "Account eliminato");
  assert.equal(priorChat.recipientId, "worker-2");
  assert.equal(priorChat.text, "Company chat should remain");
  assert.equal(priorChat.metadata.commessaId, "job-test");
  const otherChat = state.docs.find((d) => d.id === "chat-from-other");
  assert.equal(otherChat.senderId, "worker-2");
  assert.equal(otherChat.recipientId, undefined); // Deleted recipient no longer linked
  assert.equal(otherChat.senderEmail, "other@example.test");
  assert.equal(otherChat.text, "Another person's message");
  const priorAppNotice = state.docs.find((d) => d.id === "app-created-by-deleted");
  assert.equal(priorAppNotice.createdByUid, undefined);
  assert.equal(priorAppNotice.createdByEmail, undefined);
  assert.equal(priorAppNotice.createdByName, "Account eliminato");
  assert.equal(priorAppNotice.body, "Keep work information");
  const otherAppNotice = state.docs.find((d) => d.id === "app-created-by-other");
  assert.equal(otherAppNotice.createdByUid, "worker-2");
  assert.equal(otherAppNotice.createdByEmail, "other@example.test");
  // Shared notifications survive, but this UID, read-state and sender identity do not.
  const sharedNotice = state.docs.find((d) => d.id === "shared-message");
  assert.equal(sharedNotice.actorId, undefined);
  assert.equal(sharedNotice.actorName, "Account eliminato");
  assert.deepEqual(sharedNotice.recipientUserIds, ["worker-2"]);
  assert.deepEqual(sharedNotice.readBy, { "worker-2": true });
  assert.deepEqual(sharedNotice.deletedBy, {});
  assert.equal(sharedNotice.message, "Retained business notification");
  const otherNotice = state.docs.find((d) => d.id === "other-message");
  assert.equal(otherNotice.actorId, "worker-2");
  assert.deepEqual(otherNotice.recipientUserIds, ["worker-2"]);
  assert.equal(otherNotice.message, "Another user's business notification");
  assert.ok(state.removed.includes("documentReminderAcknowledgements/ack-1"));
  assert.ok(state.removed.includes("passwordRecoveryChallenges/challenge-1"));
  assert.ok(!state.removed.includes("passwordRecoveryChallenges/challenge-2"));
  for (const protectedPath of ["documents/shared-1", "documents/other-1", "platformUsers/worker-2",
    "documentReminderAcknowledgements/ack-2"]) {
    assert.ok(!state.removed.includes(protectedPath), "must preserve " + protectedPath);
  }

  // A company document must never lose a file still referenced by it.
  state = setup();
  state.docs.find((x) => x.id === "shared-1").thumbnailPath =
    "documents/worker-1/doc1/thumbnail.jpg";
  await state.handler({ auth: goodAuth, data: { confirmation: "ELIMINA" } });
  assert.deepEqual(state.removedFiles, ["documents/worker-1/doc1/private.pdf"]);
  assert.ok(state.removed.includes("documents/private-1"));
  assert.ok(!state.removed.includes("documents/shared-1"));

  state = setup();
  state.docs.find((x) => x.id === "private-1").storagePath =
    "documents/worker-2/doc3/other.pdf";
  await rejectedCode(state.handler({ auth: goodAuth, data: { confirmation: "ELIMINA" } }),
    "failed-precondition");
  assert.equal(state.removedUsers.length, 0);
  assert.equal(state.removedFiles.length, 0);
  // Older document records have no visibility field but are rendered as private.
  // Their metadata and files must be included in the deletion plan.
  state = setup();
  delete state.docs.find((x) => x.id === "private-1").visibility;
  const legacyReady = await state.handler({ auth: goodAuth, data: { confirmation: "ELIMINA", mode: "preflight" } });
  assert.equal(legacyReady.personalDocuments, 1);
  await state.handler({ auth: goodAuth, data: { confirmation: "ELIMINA" } });
  assert.ok(state.removed.includes("documents/private-1"));
  assert.equal(state.removedFiles.length, 2);

  // Legacy records lacking a valid Storage path cannot be silently orphaned.
  // Fail closed; an operator must resolve the legacy attachment before deletion.
  state = setup();
  const legacyNoStorage = state.docs.find((x) => x.id === "private-1");
  delete legacyNoStorage.visibility;
  delete legacyNoStorage.storagePath;
  await rejectedCode(state.handler({ auth: goodAuth, data: { confirmation: "ELIMINA", mode: "preflight" } }),
    "failed-precondition");
  assert.equal(state.removedUsers.length, 0);
  assert.equal(state.removed.length, 0);
  assert.equal(state.removedFiles.length, 0);

  // Transient cloud failures cannot produce a false Firebase Auth deletion.
  for (const [mode, firstPhase] of [
    ["failStorageOnce", "storage"], ["failFirestoreOnce", "firestore"],
    ["failPersonnelOnce", "personale"], ["failAuthOnce", "auth"]
  ]) {
    state = setup({ [mode]: true });
    await rejectedCode(state.handler({ auth: goodAuth, data: { confirmation: "ELIMINA" } }), "internal");
    assert.deepEqual(state.removedUsers, [], firstPhase + " unexpectedly removed Auth user");
    assert.equal(state.docs.find((d) => d.id === "employee-2").linkedUserId, "worker-2");
    const retry = await state.handler({ auth: goodAuth, data: { confirmation: "ELIMINA" } });
    assert.equal(retry.deleted, true);
    assert.deepEqual(state.removedUsers, ["worker-1"]);
    assert.equal(state.docs.find((d) => d.id === "employee-2").linkedUserId, "worker-2");
  }

  console.log("PASS: Account deletion simulated — auth, confirmation, UID isolation, retention, storage safety, failure retry.");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
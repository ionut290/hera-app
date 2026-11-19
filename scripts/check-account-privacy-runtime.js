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

function setup() {
  const users = new Set(["worker-1", "worker-2"]);
  const docs = [
    { collection: "documents", id: "private-1", ownerUserId: "worker-1", visibility: "personal", storagePath: "documents/worker-1/doc1/private.pdf", thumbnailPath: "documents/worker-1/doc1/thumbnail.jpg" },
    { collection: "documents", id: "shared-1", ownerUserId: "worker-1", visibility: "shared", storagePath: "documents/worker-1/doc2/shared.pdf" },
    { collection: "documents", id: "other-1", ownerUserId: "worker-2", visibility: "personal", storagePath: "documents/worker-2/doc3/other.pdf" },
    { collection: "documentReminderAcknowledgements", id: "ack-1", userId: "worker-1" },
    { collection: "documentReminderAcknowledgements", id: "ack-2", userId: "worker-2" },
    { collection: "platformUsers", id: "worker-1" },
    { collection: "platformUsers", id: "worker-2" },
    { collection: "privateDocuments", id: "worker-1" }
  ];
  const removed = [];
  const removedFiles = [];
  const removedUsers = [];
  const db = {
    collection(collection) {
      return {
        where(field, op, value) {
          assert.equal(op, "==");
          return {
            async get() {
              return {
                docs: docs.filter((d) => d.collection === collection && d[field] === value)
                  .map((d) => ({
                    id: d.id,
                    ref: { path: d.collection + "/" + d.id },
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
      removed.push(ref.path);
    }
  };
  const admin = {
    firestore: () => db,
    auth: () => ({
      getUser: async (uid) => {
        if (!users.has(uid)) throw new Error("not found");
        return { uid };
      },
      deleteUser: async (uid) => {
        assert.ok(users.has(uid));
        removedUsers.push(uid);
        users.delete(uid);
      }
    }),
    storage: () => ({
      bucket: () => ({
        file: (name) => ({
          async delete(options) {
            assert.equal(options.ignoreNotFound, true);
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
  return { handler, docs, removed, removedFiles, removedUsers };
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
  assert.ok(state.removed.includes("documentReminderAcknowledgements/ack-1"));
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

  console.log("PASS: Account deletion simulated — auth, confirmation, UID isolation, retention, storage safety.");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
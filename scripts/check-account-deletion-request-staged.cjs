"use strict";

// Run the REAL Cloud Function source with synthetic Firestore/Admin/Auth stubs.
// No Firebase connection, credentials, actual identities, or destructive calls.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const source = fs.readFileSync("functions/account-deletion-request.js", "utf8");
const entrypoint = fs.readFileSync("functions/main.js", "utf8");
const routineDeploy = fs.readFileSync(".github/workflows/deploy-firebase-functions.yml", "utf8");

assert.doesNotMatch(entrypoint, /require\(["']\.\/account-deletion-request["']\)/,
  "Staged function must NOT be exposed before release approval");
assert.match(routineDeploy, /!functions\/account-deletion-request\.js/,
  "Staged request handler must be excluded from automatic production deploy");
assert.doesNotMatch(routineDeploy, /functions:requestOwnVargaAccountDeletion/,
  "Staged request must not be in the normal production deploy allowlist");

class HttpsError extends Error {
  constructor(code, message) { super(message); this.code = code; }
}
const records = new Map();
const updates = [];
const admin = {
  firestore: Object.assign(() => ({
    collection(name) {
      assert.equal(name, "accountDeletionRequests");
      return {
        doc(uid) {
          return {
            uid,
            async get() {
              return { exists: records.has(uid), data: () => records.get(uid) };
            }
          };
        }
      };
    },
    async runTransaction(work) {
      return work({
        async get(ref) { return ref.get(); },
        create(ref, value) {
          if (records.has(ref.uid)) throw Error("Duplicate create");
          records.set(ref.uid, value);
          updates.push(ref.uid);
        },
        update(ref, patch) {
          if (!records.has(ref.uid)) throw Error("Missing document to update");
          records.set(ref.uid, { ...records.get(ref.uid), ...patch });
          updates.push(ref.uid);
        }
      });
    }
  }), { FieldValue: { serverTimestamp: () => "<synthetic-server-time>" } })
};
const handlers = {};
vm.runInNewContext(source, {
  exports: handlers,
  require(name) {
    if (name === "firebase-admin") return admin;
    if (name === "firebase-functions/v2/https") {
      return { onCall(opts, handler) {
        assert.equal(opts.region, "europe-west1");
        assert.equal(opts.invoker, "public");
        assert.ok(opts.maxInstances <= 3);
        return handler;
      }, HttpsError };
    }
    throw Error("Unexpected module: " + name);
  },
  Date,
  Math,
  Number,
  Object,
  Array,
  Set,
  String,
  console: { error() {} }
}, { filename: "functions/account-deletion-request.js" });

const request = handlers.requestOwnVargaAccountDeletion;
const status = handlers.getOwnVargaAccountDeletionStatus;
assert.equal(typeof request, "function");
assert.equal(typeof status, "function");

const now = Math.floor(Date.now() / 1000);
const auth = uid => ({ uid, token: { auth_time: now } });
const valid = uid => ({ auth: auth(uid), data: { confirmation: "RICHIEDO ELIMINAZIONE" } });
async function rejectsCode(promise, code) {
  await assert.rejects(promise, err => err instanceof HttpsError && err.code === code);
}
(async () => {
  await rejectsCode(request({auth:null,data:valid("worker-one").data}), "unauthenticated");
  await rejectsCode(status({auth:{uid:"worker/other"},data:{}}), "unauthenticated");
  await rejectsCode(request({auth:auth("worker-one"),data:{confirmation:"no"}}), "invalid-argument");
  await rejectsCode(request({auth:auth("worker-one"),data:{confirmation:"RICHIEDO ELIMINAZIONE",uid:"other"}}), "invalid-argument");
  await rejectsCode(request({auth:auth("worker-one"),data:{confirmation:"RICHIEDO ELIMINAZIONE",email:"pii@example.test"}}), "invalid-argument");
  await rejectsCode(request({auth:{uid:"worker-one",token:{auth_time:now-601}},data:valid("worker-one").data}), "failed-precondition");
  assert.equal(records.size, 0);

  const first = await request(valid("worker-one"));
  assert.equal(first.requested, true);
  assert.equal(first.status, "pending");
  assert.equal(first.alreadyExists, false);
  assert.deepEqual(updates, ["worker-one"]);
  assert.deepEqual(Object.keys(records.get("worker-one")).sort(),
    ["origin", "reviewBeforeErasureRequired", "status", "submittedAt", "uid"].sort());
  assert.equal(records.get("worker-one").uid, "worker-one");
  assert.equal(records.get("worker-one").reviewBeforeErasureRequired, true);
  assert.equal(records.get("worker-one").submittedAt, "<synthetic-server-time>");

  const again = await request(valid("worker-one"));
  assert.equal(again.alreadyExists, true);
  assert.equal(updates.length, 1);
  assert.equal((await status({auth:auth("worker-one"),data:{}})).status, "pending");
  assert.equal((await status({auth:auth("worker-two"),data:{}})).requested, false);
  await rejectsCode(status({auth:auth("worker-one"),data:{ uid:"worker-two" }}), "invalid-argument");

  records.get("worker-one").status = "in_review";
  assert.equal((await request(valid("worker-one"))).status, "in_review");
  assert.equal(updates.length, 1);
  records.get("worker-one").status = "completed";
  assert.equal((await request(valid("worker-one"))).status, "completed");
  assert.equal(updates.length, 1);

  // A rejected request must be appealable in-app by the same authenticated
  // account without erasing the previous review information. One retry
  // creates one new review cycle; fast consecutive taps remain idempotent.
  records.get("worker-one").status = "rejected";
  records.get("worker-one").reviewerNote = "Synthetic previous disposition";
  const reopened = await request(valid("worker-one"));
  assert.equal(reopened.status, "pending");
  assert.equal(reopened.alreadyExists, true);
  assert.equal(reopened.reopened, true);
  assert.equal(records.get("worker-one").reopenCount, 1);
  assert.equal(records.get("worker-one").reopenedAt, "<synthetic-server-time>");
  assert.equal(records.get("worker-one").reviewerNote, "Synthetic previous disposition");
  assert.equal(updates.length, 2);
  const noDuplicate = await request(valid("worker-one"));
  assert.equal(noDuplicate.status, "pending");
  assert.equal(noDuplicate.reopened, undefined);
  assert.equal(updates.length, 2);
  records.get("worker-one").status = "rejected";
  records.get("worker-one").reopenCount = "corrupt";
  await rejectsCode(request(valid("worker-one")), "failed-precondition");
  assert.equal(updates.length, 2);
  records.get("worker-one").status = "bad-state";
  await rejectsCode(request(valid("worker-one")), "failed-precondition");
  assert.equal(records.size, 1);

  const newRequest = await request(valid("worker-two"));
  assert.equal(newRequest.status, "pending");
  assert.equal(records.size, 2);
  assert.equal(records.get("worker-one").uid, "worker-one");
  console.log("PASS: staged deletion request intake is authenticated, fresh, UID-isolated, idempotent, PII-minimal and non-destructive.");
})().catch(error => { console.error(error); process.exitCode=1; });

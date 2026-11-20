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
assert.match(routineDeploy, /^on:\n  workflow_dispatch:/m);
assert.doesNotMatch(routineDeploy, /^  push:/m,
  "Automatic Firebase deploy must be disabled");
assert.match(routineDeploy, /node scripts\/check-solo-owner-release-policy\.cjs --verify/);
assert.doesNotMatch(routineDeploy, /functions:requestOwnVargaAccountDeletion/,
  "Staged request must not be in the normal production deploy allowlist");

class HttpsError extends Error {
  constructor(code, message) { super(message); this.code = code; }
}
const records = new Map();
const updates = [];
const activeUsers = new Map([
  ["worker-one", { disabled: false, tokensValidAfterTime: new Date(0).toUTCString() }],
  ["worker-two", { disabled: false, tokensValidAfterTime: new Date(0).toUTCString() }]
]);
const admin = {
  auth() {
    return { async getUser(uid) {
      if (!activeUsers.has(uid)) {
        const error = new Error("No synthetic Auth record");
        error.code = "auth/user-not-found";
        throw error;
      }
      return activeUsers.get(uid);
    }};
  },
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

  // A token for a deleted, never-created or disabled Auth user must not
  // create a request or expose status (even when its auth_time is fresh).
  await rejectsCode(request(valid("ghost-worker")), "unauthenticated");
  await rejectsCode(status({ auth: auth("ghost-worker"), data: {} }), "unauthenticated");
  activeUsers.get("worker-one").disabled = true;
  await rejectsCode(request(valid("worker-one")), "unauthenticated");
  await rejectsCode(status({ auth: auth("worker-one"), data: {} }), "unauthenticated");
  assert.equal(records.size, 0);
  activeUsers.get("worker-one").disabled = false;

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

  // A revoked ID token may still be cryptographically valid for some time.
  // Neither intake nor status may accept it once live Firebase Auth has
  // invalidated sessions, even if the user remains enabled.
  const reviewer = activeUsers.get("worker-one");
  const alreadySubmitted = updates.length;
  const oldAuth = { uid: "worker-one", token: { auth_time: now - 10 } };
  reviewer.tokensValidAfterTime = new Date((now - 5) * 1000).toUTCString();
  await rejectsCode(request({auth:oldAuth, data:valid("worker-one").data}), "unauthenticated");
  await rejectsCode(status({auth:oldAuth,data:{}}), "unauthenticated");
  await rejectsCode(status({auth:{uid:"worker-one",token:{}},data:{}}), "unauthenticated");
  await rejectsCode(status({auth:{uid:"worker-one",token:{auth_time:now+120}},data:{}}), "unauthenticated");
  assert.equal(updates.length, alreadySubmitted);
  // If live Auth revocation metadata is malformed or unavailable, do not
  // return a potentially sensitive request status or create a request.
  reviewer.tokensValidAfterTime = "invalid-auth-timestamp";
  await rejectsCode(status({auth:auth("worker-one"),data:{}}), "internal");
  await rejectsCode(request(valid("worker-one")), "internal");
  assert.equal(updates.length, alreadySubmitted);
  reviewer.tokensValidAfterTime = undefined;
  assert.equal((await status({auth:auth("worker-one"),data:{}})).status, "pending",
    "Firebase Auth may omit validSince until a first revocation");
  assert.equal((await request(valid("worker-one"))).alreadyExists, true);
  assert.equal(updates.length, alreadySubmitted);
  reviewer.tokensValidAfterTime = new Date(0).toUTCString();

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
  records.get("worker-one").reviewerNote = "Synthetic private reviewer note: NEVER return";
  records.get("worker-one").rejectionReasonCode = "retention_review_required";
  const rejection = await status({auth:auth("worker-one"),data:{}});
  assert.deepEqual(JSON.parse(JSON.stringify(rejection)),{
    requested:true,status:"rejected",reasonCode:"retention_review_required"
  });
  assert.equal(JSON.stringify(rejection).includes("private reviewer"),false);
  assert.equal((await status({auth:auth("worker-two"),data:{}})).requested,false);
  records.get("worker-one").rejectionReasonCode = "<private data not in whitelist>";
  assert.deepEqual(JSON.parse(JSON.stringify(await status({auth:auth("worker-one"),data:{}}))),{
    requested:true,status:"rejected"
  });
  records.get("worker-one").rejectionReasonCode = "retention_review_required";
  const reopened = await request(valid("worker-one"));
  assert.equal(reopened.status, "pending");
  assert.equal(reopened.alreadyExists, true);
  assert.equal(reopened.reopened, true);
  assert.equal(records.get("worker-one").reopenCount, 1);
  assert.equal(records.get("worker-one").reopenedAt, "<synthetic-server-time>");
  assert.equal(records.get("worker-one").reviewerNote, "Synthetic private reviewer note: NEVER return");
  // Historical rejection metadata remains stored for audit, but is not
  // disclosed while a renewed request is pending.
  assert.deepEqual(JSON.parse(JSON.stringify(await status({auth:auth("worker-one"),data:{}}))),{
    requested:true,status:"pending"
  });
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
  console.log("PASS: staged deletion request intake and status reject live Auth revocations, missing timestamps, disabled users and foreign UIDs; no destructive operations.");
})().catch(error => { console.error(error); process.exitCode=1; });

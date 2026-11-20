"use strict";
// Actual Apple review account creation handler with ONLY synthetic Auth and
// Firestore stubs. No Firebase SDK connection, credentials, or cloud writes.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

class HttpsError extends Error {
  constructor(code, message) { super(message); this.code = code; }
}
const primaryEmail = "ionut29019@gmail.com"; // fixed admin allowlist in source
const otherAdminEmail = "synthetic-admin@example.test";
const users = new Map([
  ["primary-admin", { uid: "primary-admin", email: primaryEmail, emailVerified: true, disabled: false, customClaims: {} }],
  ["second-admin", { uid: "second-admin", email: otherAdminEmail, emailVerified: true, disabled: false, customClaims: {} }]
]);
const byEmail = new Map([...users.values()].map(user => [user.email, user.uid]));
const customClaims = new Map();
const profiles = new Map();
let created = 0;
let allowlist = [otherAdminEmail];
let allowlistReads = 0;
const admin = {
  auth() {
    return {
      async getUser(uid) {
        const u = users.get(uid);
        if (!u) {
          const e = new Error("Synthetic missing Firebase Auth user");
          e.code = "auth/user-not-found";
          throw e;
        }
        return { ...u };
      },
      async getUserByEmail(email) {
        if (!byEmail.has(email)) {
          const e = new Error("No synthetic user");
          e.code = "auth/user-not-found";
          throw e;
        }
        return users.get(byEmail.get(email));
      },
      async createUser(payload) {
        assert.equal(payload.emailVerified, true);
        assert.equal(payload.disabled, false);
        assert.ok(payload.password.length >= 26);
        const uid = "review-" + (++created);
        const user = { uid, ...payload };
        users.set(uid, user);
        byEmail.set(payload.email, uid);
        return user;
      },
      async setCustomUserClaims(uid, claims) {
        assert.ok(users.has(uid));
        customClaims.set(uid, claims);
        users.get(uid).customClaims = claims;
      },
      async deleteUser(uid) {
        const user = users.get(uid);
        if (user) byEmail.delete(user.email);
        users.delete(uid);
      }
    };
  },
  firestore: Object.assign(() => ({
    collection(name) {
      if (name === "appConfig") {
        return { doc(id) {
          assert.equal(id, "adminUsers");
          return { async get() {
            allowlistReads++;
            return { exists: true, data: () => ({ emails: allowlist }) };
          }};
        }};
      }
      if (name === "platformUsers") {
        return { doc(uid) {
          return { async set(data) { profiles.set(uid, data); }};
        }};
      }
      throw Error("Unexpected collection: " + name);
    }
  }), { FieldValue: { serverTimestamp: () => "synthetic-time" } })
};
const exportsForTest = {};
const factory = vm.runInThisContext(
  "(function(exports, require, console) {\n" +
    fs.readFileSync("functions/apple-review-accounts.js", "utf8") + "\n})",
  { filename: "functions/apple-review-accounts.js" }
);
factory(exportsForTest, name => {
  if (name === "node:crypto") return require("node:crypto");
  if (name === "firebase-admin") return admin;
  if (name === "firebase-functions/v2/https") {
    return { HttpsError, onCall(opts, fn) {
      assert.equal(opts.region, "europe-west1");
      assert.ok(opts.maxInstances <= 2);
      return fn;
    }};
  }
  throw Error("Unexpected production dependency: " + name);
}, { error() {} });
const create = exportsForTest.createAppleReviewAccount;
assert.equal(typeof create, "function");

const now = Math.floor(Date.now() / 1000);
const auth = (uid, email, overrides = {}) => ({
  uid, token: { email, email_verified: true, auth_time: now, ...overrides }
});
const primary = () => auth("primary-admin", primaryEmail);
const request = (authContext, email = "candidate-reviewer@example.test") =>
  ({ auth: authContext, data: { email } });
async function denied(input, code) {
  const before = created, size = profiles.size;
  await assert.rejects(create(input), error =>
    error instanceof HttpsError && error.code === code);
  assert.equal(created, before, "Denied call created an Auth user");
  assert.equal(profiles.size, size, "Denied call wrote an account profile");
}

(async () => {
  await denied(request(null), "permission-denied");
  await denied(request(auth("ghost-admin", primaryEmail)), "permission-denied");
  await denied(request(auth("primary-admin", primaryEmail, { auth_time: now - 601 })),
    "failed-precondition");
  await denied(request(auth("primary-admin", primaryEmail, { auth_time: now + 61 })),
    "failed-precondition");
  await denied(request(auth("primary-admin", primaryEmail, { email_verified: false })),
    "permission-denied");
  await denied(request(auth("primary-admin", primaryEmail, { apple_review: true })),
    "permission-denied");

  users.get("primary-admin").disabled = true;
  await denied(request(primary()), "permission-denied");
  users.get("primary-admin").disabled = false;

  users.get("primary-admin").emailVerified = false;
  await denied(request(primary()), "permission-denied");
  users.get("primary-admin").emailVerified = true;

  users.get("primary-admin").email = "changed-admin@example.test";
  await denied(request(primary()), "permission-denied");
  users.get("primary-admin").email = primaryEmail;

  users.get("primary-admin").customClaims = { apple_review: true };
  await denied(request(primary()), "permission-denied");
  users.get("primary-admin").customClaims = {};

  await denied(request(auth("primary-admin", otherAdminEmail)), "permission-denied");
  await denied(request(auth("second-admin", "unknown@example.test")), "permission-denied");
  await denied(request(primary(), "bad-address"), "invalid-argument");
  assert.equal(allowlistReads, 0, "Primary admin should not require Firestore allowlist");

  allowlist = [];
  await denied(request(auth("second-admin", otherAdminEmail)), "permission-denied");
  allowlist = [otherAdminEmail];

  const result = await create(request(primary()));
  assert.equal(result.created, true);
  assert.equal(result.email, "candidate-reviewer@example.test");
  assert.ok(result.password.length >= 26);
  assert.ok(/[A-Z]/.test(result.password));
  assert.ok(/[a-z]/.test(result.password));
  assert.ok(/[0-9]/.test(result.password));
  assert.equal(created, 1);
  const uid = byEmail.get(result.email);
  assert.deepEqual(customClaims.get(uid), { apple_review: true });
  const p = profiles.get(uid);
  assert.equal(p.isAdmin, false);
  assert.equal(p.admin, false);
  assert.equal(p.role, "apple_review");
  assert.equal(p.email, result.email);
  assert.equal(p.statoAccount, "attivo");
  await denied(request(primary()), "already-exists");

  const secondary = await create(request(auth("second-admin", otherAdminEmail),
    "another-reviewer@example.test"));
  assert.equal(secondary.created, true);
  assert.equal(allowlistReads > 0, true);
  assert.equal(created, 2);
  assert.deepEqual(customClaims.get(byEmail.get(secondary.email)), { apple_review: true });

  console.log("PASS: Apple reviewer creation requires fresh active live-verified administrator; disabled, stale, renamed, unverified or reviewer admins denied without account creation.");
})().catch(error => { console.error(error); process.exitCode = 1; });

"use strict";

/*
 * IRREVERSIBLE. Deletes ONLY the exact disposable review/demo Firebase Auth
 * account, after 20 empty work-link checks and a server preflight with zero
 * personal documents. A protected, manually reviewed CI workflow is the only
 * approved entrypoint. No production business data is ever seeded or removed.
 */
const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");
const { validateTarget, assertNoWorkLinks } = require("./prepare-app-review-demo-claims.cjs");
const { assertDemoAccount, lockDown } = require("./run-approved-demo-preflight.cjs");

const PROJECT = "hera-app-6cd2b";
const ALIAS = "sqadrahera+varga-demo@gmail.com";
const EXPECTED_HASH = "458f589f8bba0bbe8dcc14d2f4e45efc5a5c3b135c2df4530413ebfdf52fa1da";
const FUNCTION_URL = "https://europe-west1-hera-app-6cd2b.cloudfunctions.net/deleteOwnVargaAccount";
const CLAIMS = Object.freeze({ appReviewDemo: true, apple_review: true });
const digest = s => crypto.createHash("sha256").update(String(s || "").trim().toLowerCase()).digest("hex");

async function request(url, data, bearer) {
  const headers = { "content-type": "application/json" };
  if (bearer) headers.authorization = "Bearer " + bearer;
  const response = await fetch(url, {
    method: "POST", headers, body: JSON.stringify(data),
    signal: AbortSignal.timeout(45000)
  });
  return { status: response.status, data: await response.json().catch(() => ({})) };
}

function validateDisposable(record) {
  validateTarget(record, ALIAS);
  assert.equal(digest(record.email), EXPECTED_HASH);
  const claims = record.customClaims || {};
  assert.ok(Object.keys(claims).length === 0 ||
    (Object.keys(claims).length === 2 && claims.apple_review === true &&
     claims.appReviewDemo === true), "Unexpected claims/roles");
  assert.ok((record.providerData || []).every(p => p.providerId === "password"),
    "Not an isolated password-only demo");
  assert.ok(Date.parse(record.metadata?.creationTime || "") >= Date.parse("2026-11-18T23:00:00Z"));
  return record.uid;
}

function webApiKey() {
  const conf = fs.readFileSync(path.resolve(__dirname, "../firebase-config.js"), "utf8");
  const m = conf.match(/\bapiKey\s*:\s*["']([^"']+)["']/);
  assert.ok(m && /^AIza[A-Za-z0-9_-]{30,}$/.test(m[1]));
  return m[1];
}

async function assertDeleted(auth, uid) {
  await assert.rejects(() => auth.getUser(uid),
    err => err?.code === "auth/user-not-found", "Target disposable Auth identity must be missing");
}

async function live() {
  assert.equal(process.env.APPROVE_DISPOSABLE_DEMO_DELETION, "1",
    "Only a human-approved protected workflow may run destructive demo test");
  const creds = JSON.parse(fs.readFileSync(process.env.GOOGLE_APPLICATION_CREDENTIALS, "utf8"));
  assert.equal(creds.project_id, PROJECT);
  const admin = require("../functions/node_modules/firebase-admin");
  const app = admin.initializeApp({ credential: admin.credential.cert(creds), projectId: PROJECT });
  let uid = null, mutationStarted = false, deletionConfirmed = false;
  try {
    const auth = admin.auth(), db = admin.firestore();
    const record = await auth.getUserByEmail(ALIAS);
    uid = validateDisposable(record);
    // The account must start DISABLED and have no unrelated Firestore links.
    assert.equal(record.disabled, true);
    assert.equal(await assertNoWorkLinks(db, uid, ALIAS), 20);
    const original = await auth.getUser(uid);
    validateDisposable(original);
    assert.equal(original.disabled, true);
    // Eliminate any reusable password while the isolated demo is disabled.
    mutationStarted = true;
    await auth.updateUser(uid, { password: crypto.randomBytes(48).toString("base64url") });
    await auth.setCustomUserClaims(uid, CLAIMS);
    await auth.revokeRefreshTokens(uid);
    assertDemoAccount(await auth.getUser(uid));
    await auth.updateUser(uid, { disabled: false });
    const custom = await auth.createCustomToken(uid);
    const login = await request(
      "https://identitytoolkit.googleapis.com/v1/accounts:signInWithCustomToken?key=" +
        encodeURIComponent(webApiKey()), { token: custom, returnSecureToken: true }
    );
    assert.equal(login.status, 200, "Isolated demo sign-in failed");
    const idToken = login.data.idToken;
    assert.ok(typeof idToken === "string" && idToken.length > 100);
    const claims = await auth.verifyIdToken(idToken);
    assert.equal(claims.uid, uid);
    assert.equal(claims.appReviewDemo, true);
    assert.equal(claims.apple_review, true);
    assert.ok(Math.floor(Date.now()/1000) - claims.auth_time < 600);

    const bad = await request(FUNCTION_URL, { data: {
      confirmation: "NON_VALIDO"
    } }, idToken);
    assert.ok(bad.status >= 400, "Invalid confirmation was accepted");

    const preflight = await request(FUNCTION_URL, { data: {
      confirmation: "ELIMINA", mode: "preflight"
    } }, idToken);
    assert.equal(preflight.status, 200, "Demo preflight failed; no deletion sent");
    assert.equal(preflight.data?.result?.ready, true);
    assert.equal(preflight.data?.result?.personalDocuments, 0,
      "Demo has personal documents and is NOT safe to delete");
    assert.equal(await assertNoWorkLinks(db, uid, ALIAS), 20);
    assert.equal((await auth.getUser(uid)).disabled, false);
    assertDemoAccount({ ...(await auth.getUser(uid)), disabled: true });

    // ONE AND ONLY ONE destructive request, targeting the authenticated
    // disposable identity. The server independently checks its email hash
    // and both claims; no user-chosen UID appears in this request.
    const erased = await request(FUNCTION_URL, { data: {
      confirmation: "ELIMINA"
    } }, idToken);
    assert.equal(erased.status, 200, "Server did not confirm deletion");
    assert.equal(erased.data?.result?.deleted, true, "No deleted=true acknowledgment");
    await assertDeleted(auth, uid);
    assert.equal(await assertNoWorkLinks(db, uid, ALIAS), 20);
    deletionConfirmed = true;
    console.log("PASS: one exact disposable demo Firebase Auth account deleted via REAL callable; 20 work-link checks still empty; no other account targeted.");
  } finally {
    if (uid && mutationStarted && !deletionConfirmed) {
      // If a network error occurred after successful deletion, no account
      // remains to lock. Otherwise immediately quarantine the demo again.
      try {
        const existing = await admin.auth().getUser(uid);
        if (digest(existing.email) === EXPECTED_HASH) {
          await lockDown(admin.auth(), uid);
        } else {
          throw Error("Cleanup identity mismatch");
        }
      } catch (error) {
        if (error?.code !== "auth/user-not-found") {
          console.error("CRITICAL: demo lockdown needs immediate manual verification.");
          throw error;
        }
      }
    }
    await app.delete();
  }
}

async function selfTest() {
  const record = { uid: "fake-disposable", email: ALIAS, disabled: true,
    metadata: { creationTime: "2026-11-19T12:00:00Z" },
    providerData: [], customClaims: {} };
  assert.equal(validateDisposable(record), "fake-disposable");
  for (const other of [
    { ...record, email: "employee@example.test" },
    { ...record, disabled: false },
    { ...record, customClaims: { admin: true } },
    { ...record, providerData: [{ providerId: "apple.com" }] },
    { ...record, metadata: { creationTime: "2024-11-19T12:00:00Z" } }
  ]) assert.throws(() => validateDisposable(other));
  await assertDeleted({ getUser: async () => {
    const e = Error("not found");e.code="auth/user-not-found";throw e;
  } }, "fake-disposable");
  await assert.rejects(() => assertDeleted({
    getUser: async () => record
  }, record.uid), /Target disposable/);
  assert.match(fs.readFileSync(__filename, "utf8"), /confirmation: "ELIMINA"/);
  console.log("PASS: irreversible demo-only deletion guard: exact email, disabled state, no Admin role, provider and creation-date checks.");
}

const flag = process.argv[2];
Promise.resolve(flag === "--self-test" ? selfTest() :
  flag === "--live" ? live() : Promise.reject(Error("Unknown test mode")))
  .catch(() => {
    console.error("Demo-only deletion E2E not confirmed. Check protected logs and safe lockdown.");
    process.exitCode=1;
  });

"use strict";

/**
 * Protected, temporary Firebase Authentication session for a DISABLED,
 * disposable app-review demo. No Firestore / Storage / Auth deletion.
 * Never display Auth tokens, passwords, email, UID, API keys or PII in logs.
 * ONLY call mode=preflight and an intentionally INVALID confirmation.
 */
const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");
const { validateTarget, assertNoWorkLinks } = require("./prepare-app-review-demo-claims.cjs");

const PROJECT = "hera-app-6cd2b";
const APPROVED_EMAIL = "sqadrahera+varga-demo@gmail.com";
const HASH = "458f589f8bba0bbe8dcc14d2f4e45efc5a5c3b135c2df4530413ebfdf52fa1da";
const ENDPOINT = "https://europe-west1-hera-app-6cd2b.cloudfunctions.net/deleteOwnVargaAccount";
const REQUIRED_CLAIMS = Object.freeze({ appReviewDemo: true, apple_review: true });

const hash = (value) => crypto.createHash("sha256").update(String(value || "").toLowerCase()).digest("hex");
function assertDemoAccount(record, { mustBeDisabled = true } = {}) {
  assert.equal(hash(record?.email), HASH, "Wrong Auth email identity");
  assert.equal(String(record?.email).toLowerCase(), APPROVED_EMAIL);
  assert.match(record?.uid || "", /^[A-Za-z0-9_-]{1,128}$/);
  if (mustBeDisabled) assert.equal(record.disabled, true, "Account must be DISABLED");
  assert.deepEqual(record.customClaims || {}, REQUIRED_CLAIMS, "Expected ONLY two demo claims");
  assert.ok((record.providerData || []).every(p => p.providerId === "password"),
    "Federated Apple/Google Auth cannot be used for disposable password demo");
  assert.ok(Date.parse(record.metadata?.creationTime || "") >= Date.parse("2026-11-18T23:00:00Z"),
    "Demo must have been created for this controlled review");
}

function getWebApiKey() {
  // The public Firebase web API key is read from our own checked-in config.
  // It is never logged or saved into an artifact.
  const config = fs.readFileSync(path.resolve(__dirname, "../firebase-config.js"), "utf8");
  const match = config.match(/\bapiKey\s*:\s*["']([^"']+)["']/);
  assert.ok(match && /^AIza[0-9a-zA-Z_-]{30,}$/.test(match[1]), "Missing web API key");
  return match[1];
}

async function postJSON(url, body, token) {
  const headers = { "content-type": "application/json" };
  if (token) headers.authorization = "Bearer " + token;
  const resp = await fetch(url, {
    method: "POST", headers, body: JSON.stringify(body),
    signal: AbortSignal.timeout(45000)
  });
  // Never print response bodies containing auth/account identifiers.
  const value = await resp.json().catch(() => ({}));
  return { status: resp.status, value };
}

async function lockDown(auth, uid) {
  const record = await auth.getUser(uid);
  assert.equal(hash(record.email), HASH, "Refuse to lock any other user");
  const errors = [];
  // Disabling FIRST is critical if any later operation fails.
  for (const [name, action] of [
    ["disable", () => auth.updateUser(uid, { disabled: true })],
    // A previously issued ID token may survive disabling; removal of
    // server-side claims makes the deployed canary deny it immediately.
    ["remove-demo-claims", () => auth.setCustomUserClaims(uid, {})],
    ["revoke-sessions", () => auth.revokeRefreshTokens(uid)]
  ]) {
    try { await action(); }
    catch (_) { errors.push(name); }
  }
  const after = await auth.getUser(uid);
  assert.equal(after.disabled, true, "Emergency lockdown did not disable demo");
  assert.deepEqual(after.customClaims || {}, {}, "Demo claims must be revoked after preflight");
  assert.equal(errors.length, 0, "Demo lockdown incomplete: " + errors.join(","));
  console.log("PASS: disposable demo DISABLED; privileged claims removed; refresh sessions revoked.");
}

async function live(mode) {
  assert.equal(process.env.APPROVE_PROTECTED_DEMO_PREFLIGHT, "1", "Protected manual workflow required");
  const filename = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  assert.ok(filename, "Protected Firebase credentials missing");
  const creds = JSON.parse(fs.readFileSync(filename, "utf8"));
  assert.equal(creds.project_id, PROJECT, "Wrong Firebase project");
  const admin = require("../functions/node_modules/firebase-admin");
  const app = admin.initializeApp({
    projectId: PROJECT, credential: admin.credential.cert(creds)
  });
  let demoUid = null;
  let changed = false;
  try {
    const auth = admin.auth();
    const record = await auth.getUserByEmail(APPROVED_EMAIL);
    demoUid = record.uid;

    if (mode === "lockdown") {
      // Cleanup cannot be blocked by the absence of claims or disabled state.
      await lockDown(auth, demoUid);
      return;
    }
    assert.equal(mode, "preflight");
    validateTarget(record, APPROVED_EMAIL);
    assertDemoAccount(record);
    assert.equal(record.metadata?.lastSignInTime || null, null,
      "Refuse to activate a demo that has already had a login");
    const db = admin.firestore();
    const checks = await assertNoWorkLinks(db, demoUid, APPROVED_EMAIL);
    assert.equal(checks, 20, "Expected exact 20 work-isolation checks");

    // Quarantine the old password before briefly enabling this exact demo.
    await auth.updateUser(demoUid, {
      password: crypto.randomBytes(48).toString("base64url")
    });
    changed = true;
    await auth.revokeRefreshTokens(demoUid);
    await auth.updateUser(demoUid, { disabled: false });

    const customToken = await auth.createCustomToken(demoUid);
    const signed = await postJSON(
      "https://identitytoolkit.googleapis.com/v1/accounts:signInWithCustomToken?key=" +
        encodeURIComponent(getWebApiKey()),
      { token: customToken, returnSecureToken: true }
    );
    assert.equal(signed.status, 200, "Temporary demo sign-in failed");
    const idToken = signed.value?.idToken;
    assert.ok(typeof idToken === "string" && idToken.length > 100,
      "Firebase Auth did not issue an ID token");
    const verified = await auth.verifyIdToken(idToken);
    assert.equal(verified.uid, demoUid);
    assert.equal(verified.appReviewDemo, true);
    assert.equal(verified.apple_review, true);
    assert.ok(Math.floor(Date.now() / 1000) - verified.auth_time < 600,
      "Fresh demo sign-in required");

    const denied = await postJSON(ENDPOINT, { data: {
      confirmation: "NON_VALIDO", mode: "preflight"
    }}, idToken);
    assert.ok(denied.status >= 400, "Invalid confirmation unexpectedly accepted");
    // Only a preflight is permitted, never the destructive request.
    const preflight = await postJSON(ENDPOINT, { data: {
      confirmation: "ELIMINA", mode: "preflight"
    }}, idToken);
    assert.equal(preflight.status, 200, "Live demo-only preflight failed");
    assert.equal(preflight.value?.result?.ready, true, "Preflight did not confirm readiness");
    assert.equal(preflight.value?.result?.personalDocuments, 0,
      "A supposedly isolated demo has personal documents; stop test");
    assert.equal((await auth.getUser(demoUid)).disabled, false);
    assert.equal(await assertNoWorkLinks(db, demoUid, APPROVED_EMAIL), 20);
    console.log("PASS: live Firebase ID token, invalid confirmation rejected, demo-only preflight ready; ZERO account deletions.");
  } finally {
    if (demoUid && (changed || mode === "lockdown")) {
      // This runs on both success and failure; a separate ALWAYS cleanup step
      // in GitHub Actions repeats it to handle runner interruptions.
      await lockDown(admin.auth(), demoUid);
    }
    await app.delete();
  }
}

async function selfTest() {
  const demo = {
    uid: "synthetic-demo", email: APPROVED_EMAIL, disabled: true,
    customClaims: { ...REQUIRED_CLAIMS },
    metadata: { creationTime: "2026-11-19T11:00:00Z" }, providerData: []
  };
  assertDemoAccount(demo);
  for (const bad of [
    { ...demo, disabled: false },
    { ...demo, email: "other@example.test" },
    { ...demo, customClaims: {} },
    { ...demo, customClaims: { appReviewDemo: true } },
    { ...demo, customClaims: { ...REQUIRED_CLAIMS, admin: true } },
    { ...demo, providerData: [{ providerId: "apple.com" }] }
  ]) assert.throws(() => assertDemoAccount(bad));
  assert.ok(!/\.deleteUser\s*\(/.test(fs.readFileSync(__filename, "utf8")));
  assert.ok(!/\bmode\s*:\s*["']delete["']/.test(fs.readFileSync(__filename, "utf8")));
  console.log("PASS: fixed demo identity, double claim and non-destructive preflight contract.");
}

const mode = process.argv[2];
if (require.main === module) {
  const p = mode === "--self-test" ? selfTest() :
    mode === "--preflight" ? live("preflight") :
    mode === "--lockdown" ? live("lockdown") :
    Promise.reject(Error("Unknown mode"));
  Promise.resolve(p).catch(() => {
    console.error("PROTECTED DEMO PREFLIGHT FAILED; no account deletion attempted.");
    process.exitCode = 1;
  });
}
module.exports = { assertDemoAccount, lockDown };

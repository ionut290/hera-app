"use strict";

// Exercises the real browser controller with synthetic Firebase/Capacitor stubs.
// No Firebase project, real Apple account or browser session is contacted.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.resolve(__dirname, "../account-privacy-controls.js"), "utf8");

function harness(options = {}) {
  const elements = new Map();
  const calls = [];
  const alerts = [];
  let confirmCalls = 0;
  let promptCalls = 0;
  let reloaded = false;
  let deleted = false;
  let signedOut = false;
  let appleSignIns = 0;
  let webAppleSignIns = 0;
  let revoked = 0;
  let checked = 0;
  let signedIn = 0;
  let requested = 0;
  let statusChecked = 0;

  class Element {
    constructor(id = "") {
      this.children = [];
      this.listeners = {};
      this.dataset = {};
      this.disabled = false;
      this.parentElement = null;
      if (id) this.id = id;
    }
    set id(value) { this._id = value; elements.set(value, this); }
    get id() { return this._id; }
    setAttribute(key, value) { this[key] = value; }
    appendChild(child) { this.children.push(child); child.parentElement = this; }
    insertAdjacentElement(_position, child) {
      if (!this.parentElement) throw Error("Missing parent for insertion");
      this.parentElement.appendChild(child);
    }
    addEventListener(name, cb) { this.listeners[name] = cb; }
    click() { return this.listeners.click?.(); }
  }
  const root = new Element("root");
  const gate = new Element("auth-gate-login-btn");
  root.appendChild(gate);
  root.appendChild(new Element("user-details-panel"));
  const menu = new Element("menu");
  root.appendChild(menu);
  const pending = new Element("access-approval-logout");
  root.appendChild(pending);
  const document = {
    readyState: "complete",
    getElementById: (id) => elements.get(id) || null,
    querySelector: (selector) => selector.includes("side-menu") ? menu : null,
    querySelectorAll: (selector) => {
      if (selector === "[data-delete-varga-account]") {
        return [...elements.values()].filter((el) => el.dataset.deleteVargaAccount);
      }
      if (selector === "[data-check-varga-deletion-status]") {
        return [...elements.values()].filter((el) => el.dataset.checkVargaDeletionStatus);
      }
      if (selector === "[data-request-varga-deletion]") {
        return [...elements.values()].filter((el) => el.dataset.requestVargaDeletion);
      }
      return [];
    },
    createElement: () => new Element(),
  };

  const user = options.noUser ? null : {
    email: "demo@example.test",
    providerData: options.appleUser ? [{ providerId: "apple.com" }] : [{ providerId: "google.com" }],
    async reauthenticateWithCredential(credential) {
      calls.push(["reauth", credential]);
      if (options.reauthFails) throw Error("Synthetic reauthentication failed");
    }
  };
  const authInstance = {
    currentUser: user,
    async signInWithCredential(credential) { signedIn++; calls.push(["firebaseAppleLogin", credential]); },
    async signInWithPopup(provider) {
      assert.ok(provider instanceof auth.OAuthProvider);
      webAppleSignIns++;
      calls.push(["webApplePopup"]);
      if (options.webPopupError) throw options.webPopupError;
      return { user };
    },
    async signOut() { signedOut = true; }
  };
  function auth() { return authInstance; }
  auth.OAuthProvider = class {
    constructor(name) { assert.equal(name, "apple.com"); }
    credential(values) { return values; }
  };

  let preflightFn = options.preflight || (async () => ({ data: { ready: true } }));
  let deleteFn = options.deleteFn || (async () => ({ data: { deleted: true } }));
  const firebase = {
    apps: [{}], auth,
    app: () => ({
      functions(region) {
        assert.equal(region, "europe-west1");
        return {
          httpsCallable(name) {
            assert.ok(["deleteOwnVargaAccount", "requestOwnVargaAccountDeletion", "getOwnVargaAccountDeletionStatus"].includes(name));
            if (name === "getOwnVargaAccountDeletionStatus") {
              return async (payload) => {
                calls.push(["statusCallable", payload]);
                statusChecked++;
                if (options.statusFn) return options.statusFn(payload);
                return { data: { requested: false } };
              };
            }
            if (name === "requestOwnVargaAccountDeletion") {
              return async (payload) => {
                calls.push(["requestCallable", payload]);
                requested++;
                if (options.requestFn) return options.requestFn(payload);
                return { data: { requested: true, status: "pending" } };
              };
            }
            return async (payload) => {
              calls.push(["callable", payload]);
              if (payload.mode === "preflight") {
                checked++;
                return preflightFn();
              }
              deleted = true;
              return deleteFn();
            };
          }
        };
      }
    })
  };
  const plugin = {
    async signInWithApple({ skipNativeAuth }) {
      assert.equal(skipNativeAuth, true);
      appleSignIns++;
      calls.push(["applePrompt"]);
      if (options.appleResponse) return options.appleResponse(appleSignIns);
      return { credential: {
        idToken: "synthetic-apple-jwt",
        nonce: "synthetic-raw-nonce",
        authorizationCode: "synthetic-authorization-code"
      } };
    },
    async revokeAccessToken({ token }) {
      assert.equal(token, "synthetic-authorization-code");
      calls.push(["revoke"]);
      if (options.revokeFails) throw Error("Synthetic Apple revocation failure");
      revoked++;
    },
    async signOut() { calls.push(["nativeSignOut"]); }
  };
  const window = {
    __VARGA_ACCOUNT_DELETION_ENABLED: options.deletionEnabled !== false,
    __VARGA_ACCOUNT_DELETION_REQUEST_ENABLED: options.requestEnabled === true,
    __VARGA_ACCOUNT_DELETION_REQUEST_SLA_DAYS: options.requestSlaDays,
    firebase,
    Capacitor: {
      isNativePlatform: () => options.platform !== "web",
      getPlatform: () => options.platform || "ios",
      Plugins: { FirebaseAuthentication: plugin }
    },
    confirm: () => { confirmCalls++; return options.confirm !== false; },
    prompt: () => { promptCalls++; return options.typedConfirmation ?? "ELIMINA"; },
    location: { reload() { reloaded = true; } }
  };
  const alert = (message) => { alerts.push(String(message)); };
  vm.runInNewContext(source, { window, document, firebase, alert,
    console: { warn() {} }, Promise }, { filename: "account-privacy-controls.js" });

  return {
    elements, alerts, calls, user, controls: {
      get confirmCalls() { return confirmCalls; },
      get promptCalls() { return promptCalls; },
      get checked() { return checked; },
      get deleted() { return deleted; },
      get signedOut() { return signedOut; },
      get reloaded() { return reloaded; },
      get appleSignIns() { return appleSignIns; },
      get webAppleSignIns() { return webAppleSignIns; },
      get revoked() { return revoked; },
      get signedIn() { return signedIn; },
      get requested() { return requested; },
      get statusChecked() { return statusChecked; }
    }
  };
}
function clickDelete(h) {
  const button = h.elements.get("varga-account-delete-menu");
  assert.ok(button, "delete account menu button missing");
  return button.click();
}
function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((yes, no) => { resolve = yes; reject = no; });
  return { promise, resolve, reject };
}
(async () => {
  let h = harness({ platform: "ios", deletionEnabled: false });
  assert.ok(h.elements.has("auth-apple-login-btn"), "Sign in with Apple must remain available");
  assert.equal(h.elements.has("varga-account-delete-menu"), false, "Deletion hidden until backend release");
  assert.equal(h.elements.has("varga-account-delete-profile"), false, "Profile deletion hidden until backend release");
  // A clear, non-destructive information entry remains visible to the user.
  assert.ok(h.elements.has("varga-deletion-info-profile"), "Hidden backend must not hide all privacy information");
  assert.ok(h.elements.has("varga-deletion-info-menu"), "Deletion status must be discoverable in normal menu");
  assert.ok(h.elements.has("varga-deletion-info-pending"), "Pending users must see deletion status");
  assert.ok(h.elements.has("varga-deletion-info-gate"), "New users must see deletion status at login gate");
  assert.equal(h.elements.get("varga-deletion-info-menu").href, "./delete-account.html",
    "Account-deletion information must link to the existing public support guide");
  assert.equal(h.alerts.length, 0,
    "Information-only link must not pretend a deletion request was sent");
  assert.equal(h.controls.checked, 0);
  assert.equal(h.controls.requested, 0);
  assert.equal(h.controls.deleted, false);
  assert.equal(h.controls.appleSignIns, 0);
  assert.equal(h.controls.signedOut, false);
  // Request intake is opt-in only; no ghost UI before backend deployment.
  // An unapproved, absent, malformed or out-of-range SLA must hide the request/status UI.
  for (const invalidDays of [undefined, null, 0, -1, 1.5, 91, "14", Infinity]) {
    const blocked = harness({ platform: "ios", deletionEnabled: false,
      requestEnabled: true, requestSlaDays: invalidDays });
    assert.equal(blocked.elements.has("varga-deletion-request-menu"), false,
      "SLA invalid: " + String(invalidDays));
    assert.equal(blocked.elements.has("varga-deletion-request-menu-status"), false);
    assert.equal(blocked.controls.requested, 0);
  }
  h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: false });
  assert.equal(h.elements.has("varga-deletion-request-menu"), false);
  h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: true, requestSlaDays: 14,
    typedConfirmation: "RICHIEDO" });
  assert.ok(h.elements.has("varga-deletion-request-menu"));
  assert.equal(h.elements.has("varga-deletion-info-menu"), false,
    "Information-only fallback must disappear once approved request backend is available");
  assert.equal(h.elements.has("varga-account-delete-menu"), false);
  await h.elements.get("varga-deletion-request-menu").click();
  assert.equal(h.controls.requested, 1);
  assert.equal(h.controls.deleted, false);
  assert.equal(h.controls.signedOut, false);
  assert.ok(h.alerts.some(text => text.includes("NON è ancora eliminato")));
  assert.ok(h.alerts.some(text => text.includes("entro 14 giorni")),
    "A signed-off SLA must be used rather than assuming a 30-day promise");
  assert.deepEqual(JSON.parse(JSON.stringify(h.calls.filter(x => x[0] === "requestCallable").map(x => x[1]))),
    [{ confirmation: "RICHIEDO ELIMINAZIONE" }]);
  h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: true, requestSlaDays: 14,
    typedConfirmation: "RICHIEDO", requestFn: async () => { throw new Error("Synthetic backend offline"); } });
  await h.elements.get("varga-deletion-request-menu").click();
  assert.equal(h.controls.requested, 1);
  assert.equal(h.controls.deleted, false);
  assert.ok(h.alerts.some(text => text.includes("non è stata confermata dal server")));
  h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: true, requestSlaDays: 14,
    typedConfirmation: "RICHIEDO", requestFn: async () => ({data:{requested:true,status:"in_review",alreadyExists:true}}) });
  await h.elements.get("varga-deletion-request-menu").click();
  assert.ok(h.alerts.some(text => text.includes("NON è ancora eliminato")));
  h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: true, requestSlaDays: 14,
    typedConfirmation: "RICHIEDO", requestFn: async () => ({data:{requested:true,status:"completed",alreadyExists:true}}) });
  await h.elements.get("varga-deletion-request-menu").click();
  assert.ok(h.alerts.some(text => text.includes("nessuna nuova cancellazione")));
  h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: true, requestSlaDays: 14,
    typedConfirmation: "NO" });
  await h.elements.get("varga-deletion-request-menu").click();
  assert.equal(h.controls.requested, 0);
  h = harness({ platform: "ios", deletionEnabled: true, requestEnabled: true, requestSlaDays: 14 });
  assert.equal(h.elements.has("varga-deletion-info-menu"), false);
  assert.equal(h.elements.has("varga-deletion-request-menu"), false);

  // A status check uses the verified caller, never creates another request.
  h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: true, requestSlaDays: 14 });
  assert.ok(h.elements.has("varga-deletion-request-menu-status"));
  await h.elements.get("varga-deletion-request-menu-status").click();
  assert.equal(h.controls.statusChecked, 1);
  assert.equal(h.controls.requested, 0);
  assert.equal(h.controls.deleted, false);
  assert.ok(h.alerts.some(x => x.includes("Nessuna richiesta")));
  assert.deepEqual(JSON.parse(JSON.stringify(
    h.calls.filter(x => x[0] === "statusCallable").map(x => x[1]))), [{}]);

  for (const [state, expected] of [
    ["pending", "in attesa di verifica"],
    ["in_review", "in revisione"],
    ["completed", "Conserva la conferma ufficiale"],
    ["rejected", "chiusa senza cancellazione"]
  ]) {
    h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: true, requestSlaDays: 14,
      statusFn: async () => ({data:{requested:true,status:state}}) });
    await h.elements.get("varga-deletion-request-menu-status").click();
    assert.ok(h.alerts.some(x => x.includes(expected)), state + " status message missing");
    assert.equal(h.controls.requested, 0);
    assert.equal(h.controls.deleted, false);
    assert.equal(h.controls.signedOut, false);
  }
  for (const data of [null, {}, {requested:true,status:"unknown"}]) {
    h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: true, requestSlaDays: 14,
      statusFn: async () => ({data}) });
    await h.elements.get("varga-deletion-request-menu-status").click();
    assert.ok(h.alerts.some(x => x.includes("Impossibile verificare lo stato")));
    assert.equal(h.controls.deleted, false);
  }
  h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: true, requestSlaDays: 14,
    noUser: true });
  await h.elements.get("varga-deletion-request-menu-status").click();
  assert.equal(h.controls.statusChecked, 0);
  assert.ok(h.alerts.some(x => x.includes("Accedi per verificare")));

  const pendingStatusResponse = deferred();
  h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: true, requestSlaDays: 14,
    statusFn: () => pendingStatusResponse.promise });
  const firstStatus = h.elements.get("varga-deletion-request-menu-status").click();
  const secondStatus = h.elements.get("varga-deletion-request-menu-status").click();
  assert.equal(h.controls.statusChecked, 1, "Duplicate concurrent status call");
  assert.equal(h.elements.get("varga-deletion-request-menu-status").disabled, true);
  pendingStatusResponse.resolve({data:{requested:true,status:"pending"}});
  await Promise.all([firstStatus, secondStatus]);
  assert.equal(h.elements.get("varga-deletion-request-menu-status").disabled, false);
  assert.equal(h.controls.deleted, false);

  h = harness({ platform: "ios", deletionEnabled: false, requestEnabled: false });
  assert.equal(h.elements.has("varga-deletion-request-menu-status"), false);
  h = harness({ platform: "ios", deletionEnabled: true, requestEnabled: true, requestSlaDays: 14 });
  assert.equal(h.elements.has("varga-deletion-request-menu-status"), false);

  h = harness({ platform: "ios" });
  assert.ok(h.elements.has("auth-apple-login-btn"), "native iOS Apple button missing");
  await h.elements.get("auth-apple-login-btn").click();
  assert.equal(h.controls.appleSignIns, 1);
  assert.equal(h.controls.signedIn, 1);
  const credential = h.calls.find((call) => call[0] === "firebaseAppleLogin")[1];
  assert.equal(credential.idToken, "synthetic-apple-jwt");
  assert.equal(credential.rawNonce, "synthetic-raw-nonce");

  h = harness({ platform: "web" });
  assert.ok(h.elements.has("auth-apple-login-btn"), "Apple OAuth button must be visible on web");
  await h.elements.get("auth-apple-login-btn").click();
  assert.equal(h.controls.webAppleSignIns, 1, "Web must use Firebase Apple OAuth popup");
  assert.equal(h.controls.appleSignIns, 0, "Web must never invoke native iOS Apple plugin");
  assert.deepEqual(h.calls.filter(x=>x[0]==="webApplePopup").map(x=>x[0]), ["webApplePopup"]);
  h = harness({ platform: "android", deletionEnabled: false });
  assert.equal(h.elements.has("auth-apple-login-btn"), false,
    "Native Android remains email/password only");
  h = harness({ platform: "web", deletionEnabled: false,
    webPopupError: {code:"auth/unauthorized-domain"} });
  await h.elements.get("auth-apple-login-btn").click();
  assert.ok(h.alerts.some(msg=>msg.includes("dominio") && msg.includes("Firebase")),
    "Unauthorized web domain must be explained without exposing internals");

  // Do not silently delete an Apple-linked Firebase account on platforms
  // without a working Apple OAuth revocation flow.
  for (const platform of ["web", "android"]) {
    h = harness({ platform, appleUser: true });
    await clickDelete(h);
    assert.equal(h.controls.checked, 1, "Preflight must still run first");
    assert.equal(h.controls.deleted, false, platform + " deleted Apple-linked account");
    assert.equal(h.controls.revoked, 0);
    assert.equal(h.controls.signedOut, false);
    assert.equal(h.controls.reloaded, false);
    assert.ok(h.alerts.some((value) =>
      value.includes("solo nell'app iPhone") &&
      value.includes("Nessun account è stato eliminato")
    ));
    // Failed attempt should release the lock and remain non-destructive.
    await clickDelete(h);
    assert.equal(h.controls.checked, 2);
    assert.equal(h.controls.deleted, false);
  }

  // Normal non-Apple users retain the existing deletion flow on the web.
  h = harness({ platform: "web" });
  await clickDelete(h);
  assert.equal(h.controls.checked, 1);
  assert.equal(h.controls.deleted, true);
  assert.equal(h.controls.signedOut, true);

  h = harness({ noUser: true });
  await clickDelete(h);
  assert.equal(h.controls.checked, 0);
  assert.equal(h.controls.deleted, false);

  h = harness({ confirm: false });
  await clickDelete(h);
  assert.equal(h.controls.checked, 0);
  assert.equal(h.controls.deleted, false);

  h = harness({ typedConfirmation: "elimina" });
  await clickDelete(h);
  assert.equal(h.controls.checked, 0);
  assert.equal(h.controls.deleted, false);

  // A denied preflight must not delete or revoke and should unlock future tries.
  let preflightAttempts = 0;
  h = harness({ preflight: async () => {
    if (++preflightAttempts === 1) throw Error("Unmapped legacy file");
    return { data: { ready: true } };
  }});
  await clickDelete(h);
  assert.equal(h.controls.deleted, false);
  assert.equal(h.controls.checked, 1);
  await clickDelete(h);
  assert.equal(h.controls.deleted, true);
  assert.equal(h.controls.checked, 2);
  assert.equal(h.controls.signedOut, true);
  assert.equal(h.controls.reloaded, true);

  // A second tap during pending preflight must never prompt or revoke twice.
  const held = deferred();
  h = harness({ preflight: () => held.promise });
  const first = clickDelete(h);
  const second = clickDelete(h);
  assert.equal(h.controls.confirmCalls, 1);
  assert.equal(h.controls.promptCalls, 1);
  assert.equal(h.controls.checked, 1);
  assert.equal(h.controls.deleted, false);
  held.resolve({ data: { ready: true } });
  await Promise.all([first, second]);
  assert.equal(h.controls.deleted, true);

  // Apple account deletion: preflight, recent Apple reauth, revoke, server delete.
  h = harness({ appleUser: true });
  await clickDelete(h);
  assert.equal(h.controls.appleSignIns, 1);
  assert.equal(h.controls.revoked, 1);
  assert.equal(h.controls.deleted, true);
  assert.deepEqual(h.calls.filter((c) => ["reauth", "revoke"].includes(c[0])).map((c) => c[0]),
    ["reauth", "revoke"]);
  assert.equal(h.controls.signedOut, true);

  // Missing fresh Apple authorizationCode must never delete or revoke.
  h = harness({ appleUser: true, appleResponse: () => ({
    credential: { idToken: "synthetic", nonce: "raw" }
  })});
  await clickDelete(h);
  assert.equal(h.controls.revoked, 0);
  assert.equal(h.controls.deleted, false);
  await clickDelete(h);
  assert.equal(h.controls.checked, 2, "revocation failure must release deletion lock");

  // If Apple revocation fails, deletion must not run at all.
  h = harness({ appleUser: true, revokeFails: true });
  await clickDelete(h);
  assert.equal(h.controls.revoked, 0);
  assert.equal(h.controls.deleted, false);
  assert.equal(h.controls.signedOut, false);
  assert.ok(h.alerts.some((x) => x.includes("Verifica Apple non riuscita")));

  // If Apple revocation succeeds BUT the Firebase deletion callable fails,
  // the account must NOT be described as deleted. Show the precise partial
  // failure because repeating Apple authorization might not be straightforward.
  h = harness({
    appleUser: true,
    deleteFn: async () => { throw Error("Synthetic Firebase cleanup error"); }
  });
  await clickDelete(h);
  assert.equal(h.controls.checked, 1);
  assert.equal(h.controls.revoked, 1);
  assert.equal(h.controls.deleted, true, "Deletion callable must be attempted");
  assert.equal(h.controls.signedOut, false);
  assert.equal(h.controls.reloaded, false);
  assert.ok(h.alerts.some((x) =>
    x.includes("Apple è stata revocata") &&
    x.includes("NON è stata confermata dal server") &&
    x.includes("contatta l’assistenza")
  ), "Specific warning must identify Apple revocation / unconfirmed Firebase deletion");
  assert.ok(!h.alerts.some((x) => x.startsWith("Account eliminato.")));

  h = harness({
    appleUser: true,
    deleteFn: async () => ({ data: { deleted: false } })
  });
  await clickDelete(h);
  assert.equal(h.controls.revoked, 1);
  assert.equal(h.controls.signedOut, false);
  assert.ok(h.alerts.some((x) => x.includes("NON è stata confermata dal server")));
  assert.ok(!h.alerts.some((x) => x.startsWith("Account eliminato.")));

  // Even a server failure must not display a false positive.
  h = harness({ deleteFn: async () => { throw Error("Temporary server failure"); } });
  await clickDelete(h);
  assert.equal(h.controls.signedOut, false);
  assert.equal(h.controls.reloaded, false);
  assert.ok(!h.alerts.some((x) => x.startsWith("Account eliminato.")));

  console.log("PASS: iOS Apple login, deletion confirmations, preflight lock, retry, revoke and success/failure paths with synthetic stubs.");
})().catch((error) => { console.error(error); process.exitCode = 1; });

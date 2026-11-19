"use strict";

// All inputs synthetic. No Firebase credentials, remote APIs, or account writes.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const root = path.resolve(__dirname, "..");
const source = fs.readFileSync(path.join(root, "apple-review-demo-deletion-ui.js"), "utf8");
const packager = fs.readFileSync(path.join(root, "scripts/prepare-capacitor-web.js"), "utf8");

// The demo UI MUST NEVER be packed into the ordinary iOS/Android app.
assert.match(packager, /excludedFilePrefixes\s*=\s*\["apple-review-"\]/);
assert.match(source, /appReviewDemo\s*===\s*true/);
assert.match(source, /apple_review\s*===\s*true/);
assert.match(source, /__VARGA_INTERNAL_DEMO_DELETION_ENABLED\s*!==\s*true/);
assert.ok(!source.includes("deleteUser("), "Client cannot delete Auth using Admin SDK");
assert.ok(!source.includes("recursiveDelete("), "Client cannot directly delete Firestore");

function harness(options = {}) {
  let user = options.noUser ? null : {
    providerData: options.appleLinked ? [{ providerId: "apple.com" }] : [{ providerId: "password" }],
    async getIdTokenResult(force) {
      assert.equal(force, true);
      return { claims: typeof options.claims === "function"
        ? options.claims() : (options.claims || {}) };
    }
  };
  const elements = new Map();
  const alerts = [];
  let listener = null;
  let preflights = 0, deletes = 0, signedOut = 0, reloaded = 0;
  class Element {
    constructor() { this.children = []; this.listeners = {}; this.disabled = false; }
    set id(value) { this._id = value; elements.set(value, this); }
    get id() { return this._id; }
    addEventListener(name, cb) { this.listeners[name] = cb; }
    appendChild(child) { this.children.push(child); child.parent = this; }
    remove() {
      if (this.parent) this.parent.children = this.parent.children.filter(x => x !== this);
      elements.delete(this.id);
    }
    click() { return this.listeners.click?.(); }
  }
  const profile = new Element();
  profile.id = "user-details-panel";
  const document = {
    readyState: "complete",
    getElementById: (id) => elements.get(id) || null,
    createElement: () => new Element()
  };
  const auth = {
    get currentUser() { return user; },
    onAuthStateChanged(cb) { listener = cb; cb(user); },
    async signOut() { signedOut++; user = null; listener?.(null); }
  };
  const firebase = {
    apps: [{}],
    auth: () => auth,
    app: () => ({
      functions(region) {
        assert.equal(region, "europe-west1");
        return {
          httpsCallable(name) {
            assert.equal(name, "deleteOwnVargaAccount");
            return async (input) => {
              assert.equal(input.confirmation, "ELIMINA");
              if (input.mode === "preflight") {
                preflights++;
                if (options.failPreflight) throw Error("synthetic preflight error");
                return { data: { ready: true } };
              }
              deletes++;
              if (options.failDelete) throw Error("synthetic deletion error");
              return { data: { deleted: true } };
            };
          }
        };
      }
    })
  };
  const window = {
    __VARGA_INTERNAL_DEMO_DELETION_ENABLED: options.enabled === true,
    firebase,
    Capacitor: { isNativePlatform: () => options.platform !== "web",
      getPlatform: () => options.platform || "ios" },
    confirm: () => options.confirm !== false,
    prompt: () => options.prompt ?? "ELIMINA",
    alert: (message) => { alerts.push(String(message)); },
    location: { reload: () => { reloaded++; } }
  };
  vm.runInNewContext(source, { window, document, Promise }, { filename: "apple-review-demo-deletion-ui.js" });
  return {
    get button() { return elements.get("varga-internal-demo-account-delete"); },
    get preflights() { return preflights; },
    get deletes() { return deletes; },
    get signedOut() { return signedOut; },
    get reloaded() { return reloaded; },
    alerts,
    switchUser(next) { user = next; listener?.(next); }
  };
}
const flush = () => new Promise(resolve => setImmediate(resolve));
const GOOD = { appReviewDemo: true, apple_review: true };

(async () => {
  for (const options of [
    { claims: GOOD }, // internal switch absent
    { enabled: true, claims: GOOD, platform: "web" },
    { enabled: true, claims: {} },
    { enabled: true, claims: { appReviewDemo: true } },
    { enabled: true, claims: { apple_review: true } },
    { enabled: true, claims: GOOD, noUser: true }
  ]) {
    const h = harness(options);
    await flush();
    assert.equal(h.button, undefined, "Non-demo/normal build must not show button");
    assert.equal(h.preflights, 0);
  }

  let claims = GOOD;
  let h = harness({ enabled: true, claims: () => claims });
  await flush();
  assert.ok(h.button, "Only verified demo must see preview");
  claims = {};
  h.button.click();
  await flush();
  assert.equal(h.deletes, 0, "Revoked claims cannot delete");
  assert.equal(h.preflights, 0, "Revoked claims cannot even preflight");

  h = harness({ enabled: true, claims: GOOD });
  await flush();
  assert.ok(h.button);
  h.switchUser(null);
  await flush();
  assert.equal(h.button, undefined, "Switching off demo hides preview");

  h = harness({ enabled: true, claims: GOOD, appleLinked: true });
  await flush();
  h.button.click();
  await flush();
  assert.equal(h.deletes, 0, "Apple-linked demo must not skip revocation");

  h = harness({ enabled: true, claims: GOOD, confirm: false });
  await flush();
  h.button.click();
  await flush();
  assert.equal(h.preflights, 0);

  h = harness({ enabled: true, claims: GOOD, failPreflight: true });
  await flush();
  h.button.click();
  await flush();
  assert.equal(h.preflights, 1);
  assert.equal(h.deletes, 0);
  assert.equal(h.reloaded, 0);

  h = harness({ enabled: true, claims: GOOD });
  await flush();
  h.button.click();
  await flush();
  assert.equal(h.preflights, 1);
  assert.equal(h.deletes, 1);
  assert.equal(h.signedOut, 1);
  assert.equal(h.reloaded, 1);
  assert.ok(h.alerts.some(x => x.includes("demo eliminato")));

  console.log("PASS: excluded demo-only preview, dual claims, logout, revoke guard, preflight and synthetic deletion.");
})().catch(e => { console.error(e); process.exitCode = 1; });

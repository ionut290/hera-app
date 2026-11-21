"use strict";

// Executes the ACTUAL service worker in an isolated VM using fake requests.
// No browser session, Firebase backend, user account or network is contacted.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.resolve(__dirname, "..");
const source = fs.readFileSync(path.join(root, "sw.js"), "utf8");
const index = fs.readFileSync(path.join(root, "index.html"), "utf8");
const privacyScript = fs.readFileSync(path.join(root, "account-privacy-controls.js"), "utf8");

const matches = [...index.matchAll(/<script\s+src="account-privacy-controls\.js\?v=([^"]+)"><\/script>/g)];
assert.equal(matches.length, 1, "Require exactly one versioned privacy client script");
assert.equal(matches[0][1], "20261121-web-apple-deletion-status1",
  "Update the privacy script revision when its revocation logic changes");
assert.match(privacyScript, /appleLinked && nativePlatform\(\) !== "ios"/,
  "Fail-closed Apple revocation protection must remain in the shipped script");
assert.match(privacyScript, /appleAuthorizationRevoked = true/,
  "Partial Apple grant revocation must be distinguishable from backend deletion");
assert.match(source, /"\/account-privacy-controls\.js",/,
  "Critical privacy script must be explicitly listed for network-first loading");

const listeners = new Map();
const cached = new Map();
let connected = true;
let currentRevision = "online-v1";
let networkCalls = 0;
const identity = request => typeof request === "string" ? request : request.url;
const syntheticResponse = revision => ({
  ok: true,
  type: "basic",
  revision,
  headers: new Headers({ "content-type": "application/javascript" }),
  clone() { return syntheticResponse(revision); }
});
const context = {
  self: {
    location: {origin: "https://pwa.invalid"},
    addEventListener(type, listener) { listeners.set(type, listener); },
    clients: {claim: async () => {}, matchAll: async () => []},
    skipWaiting() {},
  },
  caches: {
    async open() {
      return {
        async put(request, response) { cached.set(identity(request), response); },
        async addAll() {}
      };
    },
    async match(request) { return cached.get(identity(request)); },
    async keys() { return []; },
    async delete() { return true; }
  },
  async fetch() {
    networkCalls++;
    if (!connected) throw new Error("Offline (synthetic)");
    return syntheticResponse(currentRevision);
  },
  URL,
  Headers,
  AbortController,
  setTimeout,
  clearTimeout,
  Date,
  console: { log() {}, warn() {}, error() {} }
};
vm.runInNewContext(source, context, { filename: "sw.js" });
const handler = listeners.get("fetch");
assert.equal(typeof handler, "function", "Service worker fetch listener missing");

async function perform(url) {
  let promise;
  const request = {
    url,
    method: "GET",
    destination: "script",
    headers: new Headers()
  };
  handler({
    request,
    respondWith(value) { promise = value; },
    waitUntil() { /* background update not used by the critical asset */ }
  });
  assert.ok(promise, "Critical privacy script fetch must be intercepted");
  return promise;
}
(async () => {
  const url = "https://pwa.invalid/account-privacy-controls.js?v=20261121-web-apple-deletion-status1";
  assert.equal((await perform(url)).revision, "online-v1");
  assert.equal(cached.get(url).revision, "online-v1");

  // A cached prior revision must not be returned when fresh network works.
  currentRevision = "online-v2";
  assert.equal((await perform(url)).revision, "online-v2");
  assert.equal(cached.get(url).revision, "online-v2");

  // Offline users may continue with the latest cached version, not index.html
  // or a fabricated successful response.
  connected = false;
  assert.equal((await perform(url)).revision, "online-v2");
  assert.equal(networkCalls, 3);

  // Versioned paths are distinct. Never substitute a different client's
  // cached script for a first-time request while offline.
  const newUrl = "https://pwa.invalid/account-privacy-controls.js?v=20261121-newer";
  assert.equal(await perform(newUrl), undefined);

  console.log("PASS: actual service worker serves fresh privacy controls online, cached copy offline, never a mismatched version.");
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});

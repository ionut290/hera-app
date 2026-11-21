"use strict";

// Real Android runtime must not load before Firebase has verified a user.
// This test creates no account, no native permission request and no network traffic.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const source = fs.readFileSync("native-android-postlogin-loader.js", "utf8");
const original = fs.readFileSync("native-android-runtime.js", "utf8");
const index = fs.readFileSync("index.html", "utf8");

async function test(platform) {
  const loadedScripts = [];
  const authListeners = [];
  const domListeners = {};
  let currentUser = null;
  const document = {
    readyState: "loading",
    head: { appendChild(node) { loadedScripts.push(node); } },
    addEventListener(name, listener) { (domListeners[name] ||= []).push(listener); },
    createElement(tag) {
      assert.equal(tag, "script");
      return { dataset: {}, remove() {} };
    }
  };
  const auth = {
    get currentUser() { return currentUser; },
    onAuthStateChanged(callback) { authListeners.push(callback); }
  };
  const window = {
    Capacitor: {
      isNativePlatform: () => platform !== "web",
      getPlatform: () => platform
    },
    firebase: { auth: () => auth }
  };
  vm.runInNewContext(source, {window,document,console: {warn() {}}}, {
    filename: "native-android-postlogin-loader.js"
  });

  if (platform !== "android") {
    assert.equal(loadedScripts.length, 1,
      "Web/iOS must retain early WhatsApp runtime");
    assert.match(loadedScripts[0].src, /native-android-runtime\.js/);
    return;
  }
  assert.equal(loadedScripts.length, 0, "Native runtime must not load on unauthenticated Android startup");
  for (const callback of domListeners.DOMContentLoaded || []) callback();
  assert.equal(authListeners.length, 1, "One Firebase auth-state listener expected");
  authListeners.forEach(callback => callback(null));
  assert.equal(loadedScripts.length, 0, "Android sign-in gate must never load native location runtime");
  currentUser = { uid: "synthetic-uid" };
  authListeners.forEach(callback => callback(currentUser));
  assert.equal(loadedScripts.length, 1, "Native runtime must load after authenticated Firebase UID");
  assert.match(loadedScripts[0].src, /native-android-runtime\.js/);
  // Pending script injection must not duplicate on auth refresh.
  authListeners.forEach(callback => callback(currentUser));
  assert.equal(loadedScripts.length, 1);
  loadedScripts[0].onload();
  authListeners.forEach(callback => callback(currentUser));
  assert.equal(loadedScripts.length, 1);
}

(async () => {
  assert.ok(index.includes("native-android-postlogin-loader.js"),
    "index.html must load the post-login startup controller");
  assert.ok(original.includes("refreshNativeLocation();"),
    "Original protected runtime is retained unchanged and loads only after login");
  await test("android");
  await test("web");
  await test("ios");
  console.log("PASS: Android native permissions code only loads after authenticated Firebase UID.");
})().catch(error => { console.error(error); process.exitCode = 1; });

"use strict";

// Run the ACTUAL protected runtime with the Android permission gate in a VM.
// No Firebase project, real Android device, or runtime permissions are used.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

(async () => {
  const gate = fs.readFileSync("android-prelogin-permission-gate.js", "utf8");
  const runtime = fs.readFileSync("native-android-runtime.js", "utf8");
  const index = fs.readFileSync("index.html", "utf8");
  assert.match(index, /<script src="android-prelogin-permission-gate\.js[^"]*"><\/script>\s*<script src="native-android-runtime\.js[^"]*"><\/script>/,
    "Guard must load synchronously before exactly one original native runtime");
  const calls = { gpsCheck: 0, gpsPrompt: 0, geofence: 0, pushCheck: 0, pushPrompt: 0 };
  const listeners = {};
  const authListeners = [];
  let currentUser = null;
  const document = {
    readyState: "complete",
    visibilityState: "visible",
    querySelector: () => null,
    createElement: () => ({ dataset: {} }),
    head: { appendChild() {} },
    addEventListener(name, listener) { (listeners[name] ||= []).push(listener); }
  };
  const nativePlugins = {
    Geolocation: {
      async checkPermissions() { calls.gpsCheck++; return { location: "prompt", coarseLocation: "prompt" }; },
      async requestPermissions() {
        calls.gpsPrompt++;
        return { location: "denied", coarseLocation: "denied" };
      },
      async getCurrentPosition() { throw new Error("No position in this test"); }
    },
    HeraGeofence: {
      async activate() { calls.geofence++; return { active: false }; }
    },
    PushNotifications: {
      async removeAllListeners() {},
      async addListener() {},
      async createChannel() {},
      async checkPermissions() { calls.pushCheck++; return { receive: "prompt" }; },
      async requestPermissions() {
        calls.pushPrompt++;
        return { receive: "denied" };
      },
      async register() {}
    }
  };
  // Keep references to real native methods: in production Capacitor may not
  // allow replacement of its plugin proxy. The runtime itself must still be
  // fail-closed even if the early proxy guard does not intercept anything.
  const unwrappedPlugins = { ...nativePlugins };
  const auth = {
    get currentUser() { return currentUser; },
    onAuthStateChanged(callback) { authListeners.push(callback); }
  };
  const window = {
    firebase: { auth: () => auth },
    Capacitor: {
      isNativePlatform: () => true,
      getPlatform: () => "android",
      Plugins: nativePlugins
    },
    dispatchEvent() {}
  };
  const context = {
    window, document, navigator: {}, console: { warn() {} },
    localStorage: {getItem: () => null, setItem() {}},
    Map, Promise, Number, String, Object, Date, Reflect, Proxy, Error,
    setTimeout, setInterval, clearInterval, CustomEvent: class {}
  };
  vm.runInNewContext(gate, context, {filename:"android-prelogin-permission-gate.js"});
  Object.assign(nativePlugins, unwrappedPlugins); // Simulate immutable Capacitor native method proxies
  vm.runInNewContext(runtime, context, {filename:"native-android-runtime.js"});
  await new Promise(resolve => setTimeout(resolve, 20));

  assert.deepEqual(calls, {
    gpsCheck: 0, gpsPrompt: 0, geofence: 0, pushCheck: 0, pushPrompt: 0
  }, "All native GPS, background location, and notification prompts must wait for login");

  // Even a feature calling the native bridge on the login screen must be denied.
  await window.HeraNativeAndroid.refreshLocation();
  await window.HeraNativeAndroid.configureBackgroundLocation();
  await window.HeraNativeAndroid.configureNotifications();
  assert.equal(calls.gpsCheck, 0);
  assert.equal(calls.gpsPrompt, 0);
  assert.equal(calls.geofence, 0);
  assert.equal(calls.pushCheck, 0);
  assert.equal(calls.pushPrompt, 0);

  currentUser = { uid: "synthetic-uid" };
  authListeners.forEach(fn => fn(currentUser));
  await new Promise(resolve => setTimeout(resolve, 50));
  assert.equal(calls.gpsCheck, 1, "GPS starts after Firebase login");
  assert.equal(calls.gpsPrompt, 1, "Location prompt starts after Firebase login");
  assert.equal(calls.geofence, 1, "Geofence starts after Firebase login");
  assert.equal(calls.pushCheck, 1, "Push notification permission checked after login");

  currentUser = null;
  authListeners.forEach(fn => fn(null));
  for (const cb of listeners.visibilitychange || []) cb();
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(calls.gpsCheck, 1, "After logout, app resume must not obtain GPS");
  assert.equal(calls.geofence, 1, "After logout, geofence activation must be blocked");
  console.log("PASS: immutable Android runtime cannot prompt GPS or push before Firebase login.");
})().catch(error => { console.error(error); process.exitCode = 1; });

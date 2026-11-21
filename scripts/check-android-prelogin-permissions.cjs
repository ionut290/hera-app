"use strict";

// Executes the real Android startup controller with synthetic Capacitor/Firebase.
// No real permissions, device, Firebase account or network is accessed.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const source = fs.readFileSync("native-android-runtime.js", "utf8");

(async () => {
  const calls = { check: 0, request: 0, geofence: 0, push: 0, pushPrompt: 0 };
  const authHandlers = [];
  let currentUser = null;
  const auth = {
    get currentUser() { return currentUser; },
    onAuthStateChanged(callback) { authHandlers.push(callback); return () => {}; }
  };
  const events = {};
  const document = {
    readyState: "complete",
    visibilityState: "visible",
    head: { appendChild() {} },
    querySelector() { return null; },
    createElement() { return { dataset: {} }; },
    addEventListener(name, callback) { (events[name] ||= []).push(callback); }
  };
  const Geolocation = {
    async checkPermissions() { calls.check++; return { location: "prompt", coarseLocation: "prompt" }; },
    async requestPermissions() {
      calls.request++;
      return { location: "denied", coarseLocation: "denied" };
    },
    async getCurrentPosition() { throw new Error("No geolocation data should be read in this test"); },
    async watchPosition() { throw new Error("Watch must not start before explicit GPS use"); }
  };
  const PushNotifications = {
    async removeAllListeners() { calls.push++; },
    async addListener() {},
    async createChannel() {},
    async checkPermissions() { return { receive: "denied" }; },
    async requestPermissions() { calls.pushPrompt++; return { receive: "denied" }; },
    async register() {}
  };
  const HeraGeofence = {
    async activate() { calls.geofence++; return { active: false }; }
  };
  const window = {
    Capacitor: {
      isNativePlatform: () => true,
      getPlatform: () => "android",
      Plugins: { Geolocation, PushNotifications, HeraGeofence }
    },
    firebase: { auth: () => auth },
    dispatchEvent() {},
    location: { href: "capacitor://localhost" }
  };
  const navigator = {};
  const localStorage = { getItem: () => null, setItem() {} };
  vm.runInNewContext(source, {
    window, navigator, document, localStorage,
    console: { warn() {} },
    Promise, Map, Number, String, Object, Date,
    setTimeout, setInterval, clearInterval,
    CustomEvent: class {}
  }, { filename: "native-android-runtime.js" });

  await Promise.resolve();
  await window.HeraNativeAndroid.refreshLocation();
  await window.HeraNativeAndroid.configureBackgroundLocation();
  await window.HeraNativeAndroid.configureNotifications();
  for (const callback of events.visibilitychange || []) callback();
  await Promise.resolve();
  assert.deepEqual(calls, { check: 0, request: 0, geofence: 0, push: 0, pushPrompt: 0 },
    "Native Android must never ask GPS, geofence or push permissions on the unauthenticated login screen");

  currentUser = { uid: "synthetic-uid" };
  authHandlers.forEach(callback => callback(currentUser));
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(calls.check, 1, "GPS may be checked after Firebase login");
  assert.equal(calls.request, 1, "Location permission may be requested only after login");
  assert.equal(calls.geofence, 1, "Geofence may start after login");
  assert.equal(calls.push, 1, "Push notification setup may start after login");

  // Auth refresh for the same user must not trigger duplicate prompts.
  authHandlers.forEach(callback => callback(currentUser));
  await new Promise(resolve => setTimeout(resolve, 20));
  assert.equal(calls.request, 1);
  assert.equal(calls.geofence, 1);

  currentUser = null;
  authHandlers.forEach(callback => callback(null));
  for (const callback of events.visibilitychange || []) callback();
  await window.HeraNativeAndroid.refreshLocation();
  await window.HeraNativeAndroid.configureBackgroundLocation();
  await window.HeraNativeAndroid.configureNotifications();
  assert.equal(calls.request, 1, "Sign-out must not trigger GPS permission");
  assert.equal(calls.geofence, 1, "Sign-out must not activate geofence");
  console.log("PASS: Android GPS/geofence/push permissions wait for verified Firebase login.");
})().catch(error => { console.error(error); process.exitCode = 1; });

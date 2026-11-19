"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

// Exercise the actual GPS subscription implementation with a fake Firestore.
// Do not connect to or modify any Firebase project.
const app = fs.readFileSync(path.resolve(__dirname, "../app.js"), "utf8");
const start = app.indexOf("function subscribeGpsRequests() {");
const end = app.indexOf("function stopGpsRequestsSubscription()", start);
assert.ok(start >= 0 && end > start, "GPS subscription source not found");
const source = app.slice(start, end);
const listeners = [];
const rendered = [];
let unsubscribed = 0;
const ui = { gpsRequestsList: { innerHTML: "" } };
const context = {
  db: {
    collection(name) {
      assert.equal(name, "gpsUpdateRequests");
      return {
        orderBy(field, direction) {
          assert.equal(field, "createdAt");
          assert.equal(direction, "desc");
          return {
            limit(n) {
              assert.equal(n, 200);
              return {
                onSnapshot(receive, fail) {
                  listeners.push({receive, fail});
                  return () => { unsubscribed++; };
                }
              };
            }
          };
        }
      };
    }
  },
  currentUser: null,
  gpsUpdateRequests: [],
  unsubscribeGpsRequests: null,
  allowed: false,
  canManageData() { return context.allowed; },
  renderGpsRequests() { rendered.push(Array.from(context.gpsUpdateRequests, x => x.id)); },
  ui,
  console: {error() {}}
};
vm.runInNewContext(source, context, {filename: "app.js:subscribeGpsRequests"});
vm.runInNewContext("subscribeGpsRequests()", context);
assert.equal(listeners.length, 0, "Unauthenticated must not subscribe");
assert.equal(rendered.length, 1, "Unauthenticated list must be cleared");
context.currentUser = {uid: "synthetic-worker"};
context.allowed = false;
vm.runInNewContext("subscribeGpsRequests()", context);
assert.equal(listeners.length, 0, "Worker must not subscribe to all GPS coordinates");
assert.equal(context.gpsUpdateRequests.length, 0);
context.currentUser = {uid: "synthetic-admin"};
context.allowed = true;
vm.runInNewContext("subscribeGpsRequests()", context);
assert.equal(listeners.length, 1, "Admin should subscribe once");
const d = { id: "synthetic-request", data() { return {operatorId: "synthetic-worker", operatorLat: 44.1, operatorLng: 11.1}; } };
listeners[0].receive({docs: [d]});
assert.deepEqual(Array.from(context.gpsUpdateRequests, x => x.id), ["synthetic-request"]);
const previousRenderCount = rendered.length;
context.currentUser = {uid: "synthetic-worker"};
context.allowed = false;
listeners[0].receive({docs: [d]});
listeners[0].fail(new Error("synthetic stale permissions"));
assert.equal(rendered.length, previousRenderCount, "Stale admin listener must not render for workers");
assert.equal(ui.gpsRequestsList.innerHTML, "", "Stale callbacks must not show errors");
vm.runInNewContext("subscribeGpsRequests()", context);
assert.equal(unsubscribed, 1, "Switching user must unsubscribe admin listener");
assert.equal(context.gpsUpdateRequests.length, 0, "GPS data must be wiped after switch");
assert.deepEqual(Array.from(rendered.at(-1)), []);
assert.equal(listeners.length, 1);
console.log("PASS: GPS subscription admin-only, prevents stale callbacks and clears on account switch.");

"use strict";
/**
 * Static, fail-closed iOS App Review privacy regression.
 * The native iOS WKWebView must not request GPS at the login gate.
 * No Firebase, location permission, device or production data is accessed.
 */
const assert = require("node:assert/strict");
const fs = require("node:fs");
const app = fs.readFileSync("app.js", "utf8");
const invocations = [...app.matchAll(/\binitGeolocation\s*\(/g)];
assert.equal(invocations.length, 2, "Expected only the function definition and the authenticated-startup invocation");
assert.doesNotMatch(app, /initHoursPage\(\);\s*initGeolocation\(\);/, "Location must not be requested on an unauthenticated startup");
assert.match(app, /if \(loggedIn\) \{[\s\S]{0,1800}runDeferredStartupTasks\(\[([\s\S]{0,220})\(\) => initGeolocation\(\)/, "GPS should start only in the authenticated deferred tasks");
assert.match(app, /function initGeolocation\(\) \{[\s\S]{0,900}navigator\.geolocation\.getCurrentPosition/, "Explicit GPS implementation must remain intact for signed-in features");
console.log("PASS: login gate has no automatic GPS prompt; authenticated GPS startup retained.");

"use strict";
/**
 * Static, fail-closed iOS App Review privacy regression.
 * The native iOS WKWebView must not request GPS at the login gate.
 * No Firebase, location permission, device or production data is accessed.
 */
const assert = require("node:assert/strict");
const fs = require("node:fs");
const app = fs.readFileSync("app.js", "utf8");
const invocations = app.split("\n").filter(line => /\binitGeolocation\s*\(/.test(line) && !line.trim().startsWith("//"));
assert.equal(invocations.length, 2, "Expected only the function definition and the authenticated-startup invocation");
assert.doesNotMatch(app, /initHoursPage\(\);\s*initGeolocation\(\);/, "Location must not be requested on an unauthenticated startup");
assert.match(app, /if \(loggedIn\) \{[\s\S]{0,1800}runDeferredStartupTasks\(\[([\s\S]{0,220})\(\) => initGeolocation\(\)/, "GPS should start only in the authenticated deferred tasks");
assert.match(app, /function initGeolocation\(\) \{[\s\S]{0,3000}navigator\.geolocation\.getCurrentPosition/, "Explicit GPS implementation must remain intact for signed-in features");
const loginHint = "Accedi con email e password oppure scegli uno dei metodi di accesso disponibili.";
assert.ok(app.includes(loginHint), "Authentication gate must not misleadingly require Google when Apple login exists");
assert.ok(fs.readFileSync("index.html", "utf8").includes(loginHint), "Static fallback login hint must match JS and include all options");
console.log("PASS: login gate has no automatic GPS prompt; authenticated GPS startup retained.");

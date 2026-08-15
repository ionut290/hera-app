#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const read = (path) => fs.readFileSync(path, "utf8");
const index = read("index.html");
const authFix = read("auth-login-fix.js");
const loginRetry = read("login-retry-fix.js");
const sw = read("sw.js");
const headerRuntime = read("header-menu-runtime.js");
const netlifyHeaders = read("_headers");

const authScriptMatch = index.match(/<script\s+src="auth-login-fix\.js[^\"]*"><\/script>/);
const appScriptMatch = index.match(/<script\s+src="app\.js[^\"]*"><\/script>/);
assert.ok(authScriptMatch, "auth-login-fix.js deve essere caricato da index.html");
assert.ok(appScriptMatch, "app.js deve essere caricato da index.html");
const authIndex = index.indexOf(authScriptMatch[0]);
const appIndex = index.indexOf(appScriptMatch[0]);
assert.ok(authIndex < appIndex, "Il controller auth deve partire prima di app.js");

assert.match(authFix, /installAuthStartupController/);
assert.match(authFix, /__heraSavedCredentialsAutoLoginInstalled\s*=\s*true/);
assert.match(authFix, /Auth\?\.Persistence\?\.LOCAL|Auth\.Persistence\.LOCAL/);
assert.match(authFix, /onIdTokenChanged/);
assert.match(authFix, /MutationObserver/);
assert.match(authFix, /if \(!authResolved\)[\s\S]*applyGateHidden\(true\)/);
assert.match(authFix, /isUsableAuthenticatedUser/);
assert.match(authFix, /visibilitychange/);
assert.match(authFix, /pageshow/);
assert.match(authFix, /HeraAuthStartupController/);

const persistenceIndex = loginRetry.indexOf("setPersistence(firebase.auth.Auth.Persistence.LOCAL)");
const signInIndex = loginRetry.indexOf("signInWithEmailAndPassword");
assert.ok(persistenceIndex >= 0 && signInIndex >= 0 && persistenceIndex < signInIndex,
  "La persistenza LOCAL deve essere impostata prima del login email/password");

assert.match(sw, /varga-cantieri-shell-v132/);
assert.match(sw, /CACHE_RESET_VERSION = "20260814-loading-humor1"/);
for (const path of [
  "/firebase-config.js",
  "/auth-login-fix.js",
  "/login-retry-fix.js",
  "/first-login-password.js",
  "/approval-access.js",
  "/header-menu-runtime.js"
]) {
  assert.ok(sw.includes(`"${path}"`), `${path} deve essere network-first`);
}
assert.doesNotMatch(sw, /header-menu-runtime-original\.js/);
assert.doesNotMatch(headerRuntime, /header-menu-runtime-original\.js/);
assert.doesNotMatch(headerRuntime, /commessa-listener-cleanup\.js/);
assert.match(headerRuntime, /__heraCommessaListenerCleanupInstalled/);
assert.match(headerRuntime, /setupHomeHeader/);
assert.doesNotMatch(netlifyHeaders, /header-menu-runtime-original\.js/);
assert.equal(fs.existsSync("header-menu-runtime-original.js"), false,
  "Il vecchio runtime header originale deve essere eliminato dal repository");
assert.equal(fs.existsSync("commessa-listener-cleanup.js"), false,
  "La pulizia listener separata deve essere incorporata nel runtime header");
assert.doesNotMatch(sw, /auto-login-saved-credentials\.js/);
assert.doesNotMatch(netlifyHeaders, /auto-login-saved-credentials\.js/);
assert.equal(fs.existsSync("auto-login-saved-credentials.js"), false,
  "Il vecchio controller auto-login deve essere eliminato dal repository");
assert.match(sw, /NETWORK_FIRST_ASSET_PATHS\.has\(url\.pathname\)/);
assert.match(sw, /networkFirstForCriticalAsset/);
assert.match(authFix, /__heraSavedCredentialsAutoLoginInstalled/);

console.log("Auth startup flow checks passed: auth unico, header runtime consolidato, persistence LOCAL e asset critici network-first.");

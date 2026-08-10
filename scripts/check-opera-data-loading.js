"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");
const updater = fs.readFileSync("pwa-login-force-update.js", "utf8");
const sharedViews = fs.readFileSync("shared-static-views.js", "utf8");
const startupOptimizer = fs.readFileSync("firestore-startup-cost-optimizer.js", "utf8");

assert.match(app, /isOperaBrowser[\s\S]*?db\.settings\(\{ experimentalForceLongPolling: true \}\)/);
assert.doesNotMatch(app, /const query = db\.collection\(commesseCollectionName\)\.orderBy\("createdAt", "desc"\)/);
assert.match(app, /const query = db\.collection\(commesseCollectionName\);/);
assert.match(sharedViews, /IS_OPERA_BROWSER[\s\S]*?reference\.get\(\)/);
assert.match(startupOptimizer, /IS_OPERA_BROWSER[\s\S]*?query\.get\(\)/);
assert.match(sw, /varga-cantieri-shell-v124/);
assert.match(updater, /const APP_VERSION = "v124"/);

console.log("Opera Firestore transport and legacy commesse loading checks passed.");

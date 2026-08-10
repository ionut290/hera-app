"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const sw = fs.readFileSync("sw.js", "utf8");
const updater = fs.readFileSync("pwa-login-force-update.js", "utf8");

assert.match(app, /isOperaBrowser[\s\S]*?db\.settings\(\{ experimentalForceLongPolling: true \}\)/);
assert.doesNotMatch(app, /const query = db\.collection\(commesseCollectionName\)\.orderBy\("createdAt", "desc"\)/);
assert.match(app, /const query = db\.collection\(commesseCollectionName\);/);
assert.match(sw, /varga-cantieri-shell-v123/);
assert.match(updater, /const APP_VERSION = "v123"/);

console.log("Opera Firestore transport and legacy commesse loading checks passed.");

#!/usr/bin/env node
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const rules = fs.readFileSync("firestore.rules", "utf8");
const privacy = fs.readFileSync("privacy.html", "utf8");

assert.match(rules, /match \/appConfig\/adminUsers\s*\{\s*allow read, write:\s*if isAdmin\(\);\s*\}/);
assert.match(rules, /match \/appConfig\/\{configId\}\s*\{[\s\S]*?configId in \["adminUsers", "driveAdminSecret", "driveBridge"\]/);
assert.match(rules, /match \/operatorPositions\/\{userId\}\s*\{\s*allow read, create, update, delete:\s*if false;/);
assert.match(privacy, /ionut29019@gmail\.com/);
assert.match(privacy, /3892352575/);
assert.match(privacy, /Hera/);
assert.match(privacy, /90 giorni/);
assert.match(privacy, /24 mesi/);
assert.match(privacy, /10 anni/);
console.log("Security policy and privacy content regression checks passed.");

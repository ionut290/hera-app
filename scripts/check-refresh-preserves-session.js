#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("update-app-feature.js", "utf8");

assert.match(source, /APP_CACHE_PREFIXES/);
assert.match(source, /varga-cantieri-shell-/);
assert.match(source, /hera-app-shell-/);
assert.match(source, /cacheNames\.filter\(isAppShellCache\)/);
assert.doesNotMatch(source, /firebase\.auth\(\)\.signOut\(/);
assert.doesNotMatch(source, /localStorage\.clear\(/);
assert.doesNotMatch(source, /sessionStorage\.clear\(/);
assert.doesNotMatch(source, /indexedDB\.deleteDatabase\(/);

console.log("Refresh cache preserves Firebase session and local app data.");

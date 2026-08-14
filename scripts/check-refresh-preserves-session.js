#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("update-app-feature.js", "utf8");
const executable = source.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");

assert.match(source, /APP_CACHE_PREFIXES/);
assert.match(source, /varga-cantieri-shell-/);
assert.match(source, /hera-app-shell-/);
assert.match(source, /cacheNames\.filter\(isAppShellCache\)/);
assert.doesNotMatch(executable, /firebase\.auth\(\)\.signOut\(/);
assert.doesNotMatch(executable, /localStorage\.clear\(/);
assert.doesNotMatch(executable, /sessionStorage\.clear\(/);
assert.doesNotMatch(executable, /indexedDB\.deleteDatabase\(/);

console.log("Refresh cache preserves Firebase session and local app data.");

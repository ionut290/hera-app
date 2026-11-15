#!/usr/bin/env node
"use strict";

// App Store Review Guideline 2.5.19: ship exactly the normal user experience.
// This audit runs after preparing www/, before any iOS or Android native build.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const root = path.resolve(__dirname, "..");
const config = JSON.parse(fs.readFileSync(path.join(root, "capacitor.config.json"), "utf8"));
const bundle = path.join(root, config.webDir || "www");
assert.ok(fs.existsSync(path.join(bundle, "index.html")), "Native web bundle is missing index.html");
assert.ok(fs.existsSync(path.join(bundle, "app.js")), "Native app functionality missing");

const forbidden = [
  /apple-review-/i,
  /\bapple_review\b/i,
  /\bappleReviewDemo\b/,
  /\bcreateAppleReviewAccount\b/,
  /\bVargaReviewFirebase\b/,
  /\breviewerAuth\b/,
  /\bassertNotAppleReviewer\b/,
  /ACCESSO REVISORI APPLE/i,
  /Esplora demo · dati fittizi/i
];
const files = [];
function walk(dir) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const item = path.join(dir, entry.name);
    const relative = path.relative(bundle, item).replace(/\\/g, "/");
    if (entry.isDirectory()) { walk(item); continue; }
    assert.ok(!/apple-review-/i.test(relative), "Reviewer-only file in native bundle: " + relative);
    if (/\.(?:js|html|css|json|webmanifest)$/i.test(relative)) {
      const content = fs.readFileSync(item, "utf8");
      for (const marker of forbidden) {
        assert.doesNotMatch(content, marker, "Reviewer-only behavior in native asset: " + relative);
      }
    }
    files.push(relative);
  }
}
walk(bundle);
assert.ok(files.length > 10, "Native assets were unexpectedly empty");
console.log("Guideline 2.5.19 audit PASSED: no Apple reviewer-only pages, demo or role-specific code in native assets (" + files.length + " files).");

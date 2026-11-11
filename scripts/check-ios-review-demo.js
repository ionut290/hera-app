#!/usr/bin/env node
"use strict";

// Safety gate for the App Store reviewer's isolated, synthetic demo.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const { execFileSync } = require("node:child_process");

const read = (name) => fs.readFileSync(name, "utf8");
const index = read("index.html");
const html = read("apple-review-demo.html");
const js = read("apple-review-demo.js");
const css = read("apple-review-demo.css");

const authStart = index.indexOf('<section id="auth-gate"');
const authEnd = index.indexOf("</section>", authStart);
const entry = index.indexOf('id="apple-review-demo-link"');
assert.ok(authStart >= 0 && authEnd > authStart && entry > authStart && entry < authEnd,
  "The demo must be linked only from the existing login gate.");
assert.ok(index.includes('href="./apple-review-demo.html"'), "The local demo link is missing.");
assert.ok(html.includes("connect-src 'none'") && html.includes("form-action 'none'") &&
  html.includes("default-src 'none'"), "Demo content security policy must prevent data transmission.");
assert.ok(html.includes('href="./index.html"'), "Reviewers must be able to return to the real login.");
assert.ok(html.includes('src="./apple-review-demo.js"') &&
  html.includes('href="./apple-review-demo.css"'), "Demo must use local assets.");
assert.doesNotMatch(html, /<script(?![^>]*src="\.\/apple-review-demo\.js")[^>]*>/i,
  "No extra scripts may run inside the isolated demo.");
assert.doesNotMatch(html, /https?:\/\//i, "Demo must not load remote assets or services.");
assert.ok(html.includes("MODALITÀ DEMO") && html.includes("DATI INVENTATI"),
  "Fictional/demo labelling must remain visible.");

// Ban APIs that could access production systems, leak data, or persist reviewer state.
const runtime = js.replace(/\/\*[\s\S]*?\*\//g, "").replace(/^\s*\/\/.*$/gm, "");
for (const [name, pattern] of Object.entries({
  "network": /\b(?:fetch|XMLHttpRequest|WebSocket|EventSource|sendBeacon)\b/i,
  "cloud services": /\b(?:firebase|firestore|httpsCallable|GoogleService|Capacitor)\b/i,
  "device location": /\b(?:geolocation|watchPosition|getCurrentPosition)\b/i,
  "persistent storage": /\b(?:localStorage|sessionStorage|indexedDB|cookie|caches)\b/i,
  "external redirects": /\b(?:window\.open|location\.assign|location\.replace)\b/i,
  "cross-context messaging": /\b(?:postMessage|BroadcastChannel)\b/i
})) assert.doesNotMatch(runtime, pattern, "Isolated demo cannot use " + name);
for (const label of ["DEMO-C01", "DEMO-C02", "DEMO-101", "DEMO-201", "Simula intervento"]) {
  assert.ok(js.includes(label) || html.includes(label), "Synthetic demo fixture missing: " + label);
}
assert.ok(css.includes("@media(max-width:700px)"), "Mobile layout is required.");
execFileSync(process.execPath, ["--check", "apple-review-demo.js"], { stdio: "inherit" });

// This script may be executed after 'npm run android:aab:prepare' as well.
if (fs.existsSync("www/index.html")) {
  for (const name of ["index.html", "apple-review-demo.html", "apple-review-demo.js", "apple-review-demo.css"]) {
    assert.ok(fs.existsSync("www/" + name), "Missing bundled iOS/Android asset " + name);
  }
  assert.ok(read("www/index.html").includes('href="./apple-review-demo.html"'),
    "The native web bundle must expose the demo via login.");
}
console.log("Apple Review demo isolation passed: synthetic data, CSP, no real services or persistent state.");

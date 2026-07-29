#!/usr/bin/env node
"use strict";

const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const browser = fs.readFileSync(path.join(root, "google-sheet-two-way-sync.js"), "utf8");
const proxy = fs.readFileSync(path.join(root, "netlify/functions/google-sheet-sync.js"), "utf8");
const auth = fs.readFileSync(path.join(root, "netlify/functions/_firebase-token.js"), "utf8");
const appsScript = fs.readFileSync(path.join(root, "google-apps-script/google-sheet-sync/Code.gs"), "utf8");
const netlify = fs.readFileSync(path.join(root, "netlify.toml"), "utf8");
const firebaseConfig = fs.readFileSync(path.join(root, "firebase-config.js"), "utf8");

const checks = [
  ["UI bidirezionale", browser.includes("Aggiorna app dal foglio") && browser.includes("Invia app al foglio")],
  ["chiavi tecniche anti duplicati", browser.includes("SYNC_KEY") && browser.includes("IMPIANTO_KEY")],
  ["push automatico Firestore", browser.includes("onSnapshot") && browser.includes("schedulePush")],
  ["pull periodico GViz", browser.includes("/api/google-sheet-import") && browser.includes("setInterval")],
  ["scrittura autenticata", browser.includes("getIdToken") && proxy.includes("authenticateEvent")],
  ["token Firebase verificato", auth.includes("RSA-SHA256") && auth.includes("securetoken.google.com")],
  ["Apps Script doPost", appsScript.includes("function doPost") && appsScript.includes("replaceRows")],
  ["segreto server side", proxy.includes("GOOGLE_SHEET_SYNC_SECRET") && appsScript.includes("SYNC_SECRET")],
  ["rotta Netlify", netlify.includes('/api/google-sheet-sync')],
  ["modulo caricato", firebaseConfig.includes("google-sheet-two-way-sync.js")]
];

let failed = false;
for (const [name, passed] of checks) {
  console.log(`${passed ? "OK" : "FAIL"} ${name}`);
  failed ||= !passed;
}
if (failed) process.exit(1);

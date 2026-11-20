"use strict";

// Read-only contract to prevent publishing the entire repository as a Firebase
// Hosting site. No network, secrets, or Firebase CLI involved.
// A deny-list alone is NOT evidence that production Hosting is release-ready.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const root = path.resolve(__dirname, "..");
const config = JSON.parse(fs.readFileSync(path.join(root, "firebase.json"), "utf8"));
const hosting = config.hosting;
assert.ok(hosting && typeof hosting === "object");
assert.equal(hosting.public, ".", "Reassess this protection if Hosting source changes");
const ignores = hosting.ignore || [];
for (const name of ["firebase.json",".firebaserc","README.md","**/.*","**/node_modules/**","functions/**","docs/**","scripts/**","android/**","ios/**","netlify/**","tests/**","test/**","coverage/**","artifacts/**","**/*.md","**/*.cjs","**/*.mjs","**/*.sh","**/*.py","**/*.rb","**/*.key","**/*.pem","**/*.p12","**/*.p8","**/*.mobileprovision","**/*.keystore","**/*.zip","**/*.tar","**/*.gz","**/*.db","**/*.sqlite","**/*.sql","**/*.csv","**/*.xlsx","**/*.xls","**/*.pdf","**/*.map","package.json","package-lock.json","netlify.toml","firestore.rules","storage.rules","firestore.indexes.json","capacitor.config.*","GoogleService-Info.plist","google-services.json"]) {
  assert.ok(ignores.includes(name), "Unsafe Hosting content: missing ignore " + name);
}
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const sw = fs.readFileSync(path.join(root, "sw.js"), "utf8");
const criticalAssets = ["index.html", "sw.js", "app.js", "account-privacy-controls.js", "delete-account.html"];
for (const asset of criticalAssets) {
  assert.ok(fs.existsSync(path.join(root, asset)), "Essential web asset absent: " + asset);
  assert.ok(!ignores.includes(asset), "Essential runtime asset excluded: " + asset);
}
for (const content of [html, sw]) {
  // No page/service-worker dependency may point into an intentionally
  // excluded private toolchain directory.
  assert.doesNotMatch(content, /(?:src|href)=["'](?:scripts|docs|netlify|android|ios)[/]/i);
  assert.doesNotMatch(content, /["'][/](?:scripts|docs|netlify|android|ios)[/]/i);
}
assert.ok(ignores.includes("**/node_modules/**"));
assert.ok(ignores.includes("**/.*"));
console.log("PASS: Firebase Hosting excludes internal source, CI, docs, secrets, archives, build tooling; core client assets are present. Manual live-site audit still required.");

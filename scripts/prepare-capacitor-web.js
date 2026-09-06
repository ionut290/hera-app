"use strict";

const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const output = path.join(root, "www");
const extensions = new Set([".css", ".html", ".js", ".json", ".webmanifest"]);
const excludedDirectories = new Set(["android", "android-resources", "docs", "functions", "netlify", "node_modules", "scripts", "tools", "www", ".git", ".github"]);

fs.rmSync(output, { recursive: true, force: true });
fs.mkdirSync(output, { recursive: true });

for (const entry of fs.readdirSync(root, { withFileTypes: true })) {
  if (!entry.isFile() || !extensions.has(path.extname(entry.name))) continue;
  fs.copyFileSync(path.join(root, entry.name), path.join(output, entry.name));
}

for (const directory of ["icons"]) {
  fs.cpSync(path.join(root, directory), path.join(output, directory), { recursive: true });
}

for (const directory of excludedDirectories) {
  if (fs.existsSync(path.join(output, directory))) {
    throw new Error(`La directory esclusa ${directory} non deve essere inclusa nel bundle web.`);
  }
}

const requiredFiles = [
  "index.html",
  "app.js",
  "green-assistant.js",
  "green-assistant.css",
  "android-whazzup-photo-order.js",
  "style.css",
  "firebase-config.js",
  "auth-login-fix.js",
  "approval-access.js",
  "approval-access.css",
  "first-login-password.css",
  "first-login-password.js",
  "password-recovery-code.js",
  "login-retry-fix.css",
  "login-retry-fix.js",
  "documents.js",
  "private-documents-v2.css",
  "identity-card-feature.js",
  "identity-card-feature.css",
  "native-android-runtime.js",
  "notification-session-enhancements.js",
  "squadre-restyle.css",
  "calendar-feature.css",
  "administrative-calendar.js",
  "today-summary-interactions.js",
  "update-app-feature.js",
  "sw.js",
  "icons/varga-cantieri-512.png"
];
for (const relativePath of requiredFiles) {
  const target = path.join(output, relativePath);
  if (!fs.existsSync(target) || fs.statSync(target).size === 0) {
    throw new Error(`Asset Android obbligatorio mancante: ${relativePath}`);
  }
}

console.log(`Bundle web Android preparato in www (${fs.readdirSync(output).length} elementi principali).`);

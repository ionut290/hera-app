#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const {
  resolveBuildVersion,
  stampServiceWorker,
  stampIndex
} = require("./stamp-pwa-build-version");

const ROOT = path.resolve(__dirname, "..");
const sw = fs.readFileSync(path.join(ROOT, "sw.js"), "utf8");
const index = fs.readFileSync(path.join(ROOT, "index.html"), "utf8");
const netlify = fs.readFileSync(path.join(ROOT, "netlify.toml"), "utf8");
const workflow = fs.readFileSync(path.join(ROOT, ".github/workflows/check-critical-flows.yml"), "utf8");

const commit = "ABCDEF1234567890ABCDEF1234567890ABCDEF12";
const version = resolveBuildVersion({ COMMIT_REF: commit });
if (version !== "git-abcdef123456") throw new Error(`Versione commit inattesa: ${version}`);

const stampedSw = stampServiceWorker(sw, version);
if (!stampedSw.includes(`const CACHE_NAME = "varga-cantieri-shell-${version}";`)) {
  throw new Error("CACHE_NAME non viene legato al commit di deploy");
}
if (!stampedSw.includes(`const CACHE_RESET_VERSION = "${version}";`)) {
  throw new Error("CACHE_RESET_VERSION non viene legato al commit di deploy");
}

const stampedIndex = stampIndex(index, version);
if (!stampedIndex.includes(`navigator.serviceWorker.register("./sw.js?v=${version}"`)) {
  throw new Error("La registrazione del Service Worker non riceve la versione build");
}

const indexAppVersion = index.match(/app\.js\?v=([^"]+)/)?.[1] || "";
const serviceWorkerAppVersion = sw.match(/app\.js\?v=([^"]+)/)?.[1] || "";
if (!indexAppVersion || indexAppVersion !== serviceWorkerAppVersion) {
  throw new Error("La versione cache di app.js deve coincidere tra index.html e sw.js");
}
for (const criticalPath of ["/app.js", "/fatto-button-immediate.js"]) {
  if (!sw.includes(`  "${criticalPath}",`)) {
    throw new Error(`${criticalPath} deve usare network-first: FATTO/FORZA non possono avviarsi con logica obsoleta`);
  }
}

// Stamp must remain FIRST even when the web-only deploy adds packaging and
// checks afterwards; otherwise Netlify would serve a stale service-worker cache.
const netlifyBuildCommand = netlify.match(/^\s*command\s*=\s*"([^"]+)"/m)?.[1] || "";
if (!netlify.includes("[build]") ||
    !netlifyBuildCommand.startsWith("node scripts/stamp-pwa-build-version.js") ||
    !netlifyBuildCommand.includes("&& node scripts/prepare-capacitor-web.js") ||
    !netlifyBuildCommand.includes("&& node scripts/prepare-netlify-web.cjs") ||
    !netlifyBuildCommand.includes("&& node scripts/check-netlify-public-build.cjs")) {
  throw new Error("Netlify deve versionare la PWA, creare il bundle web isolato e validarlo");
}
if (!workflow.includes("node scripts/check-pwa-build-versioning.js")) {
  throw new Error("Il controllo PWA non è collegato alla CI critica");
}
if (!index.includes("PWA_EMERGENCY_CACHE_RESET_VERSION")) {
  throw new Error("Il reset cache d'emergenza separato è stato rimosso");
}

console.log("PWA build versioning checks passed.");
console.log(`Deploy di prova -> ${version}`);

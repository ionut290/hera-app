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

if (!/\[build\][\s\S]*command\s*=\s*["']node scripts\/stamp-pwa-build-version\.js["']/.test(netlify)) {
  throw new Error("Netlify non esegue lo stamp automatico della versione PWA");
}
if (!workflow.includes("node scripts/check-pwa-build-versioning.js")) {
  throw new Error("Il controllo PWA non è collegato alla CI critica");
}
if (!index.includes("PWA_EMERGENCY_CACHE_RESET_VERSION")) {
  throw new Error("Il reset cache d'emergenza separato è stato rimosso");
}

console.log("PWA build versioning checks passed.");
console.log(`Deploy di prova -> ${version}`);

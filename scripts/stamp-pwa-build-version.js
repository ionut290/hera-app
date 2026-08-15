#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.resolve(__dirname, "..");
const SW_PATH = path.join(ROOT, "sw.js");
const INDEX_PATH = path.join(ROOT, "index.html");

function normalizeBuildVersion(value) {
  const raw = String(value || "").trim();
  if (!raw) return "local-static";
  const safe = raw.toLowerCase().replace(/[^a-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "");
  if (!safe) return "local-static";
  return safe.length > 24 ? safe.slice(0, 24) : safe;
}

function resolveBuildVersion(env = process.env) {
  const commit = String(env.COMMIT_REF || "").trim();
  if (commit) return `git-${normalizeBuildVersion(commit).slice(0, 12)}`;
  const deploy = String(env.DEPLOY_ID || "").trim();
  if (deploy) return `deploy-${normalizeBuildVersion(deploy).slice(0, 16)}`;
  return normalizeBuildVersion(env.VARGA_BUILD_VERSION || "local-static");
}

function stampServiceWorker(source, version) {
  let next = String(source);
  const cacheNamePattern = /const CACHE_NAME = ["'][^"']+["'];/;
  const resetVersionPattern = /const CACHE_RESET_VERSION = ["'][^"']+["'];/;
  if (!cacheNamePattern.test(next)) throw new Error("CACHE_NAME non trovato in sw.js");
  if (!resetVersionPattern.test(next)) throw new Error("CACHE_RESET_VERSION non trovato in sw.js");
  next = next.replace(cacheNamePattern, `const CACHE_NAME = "varga-cantieri-shell-${version}";`);
  next = next.replace(resetVersionPattern, `const CACHE_RESET_VERSION = "${version}";`);
  return next;
}

function stampIndex(source, version) {
  const pattern = /navigator\.serviceWorker\.register\(["']\.\/sw\.js\?v=[^"']+["']/;
  if (!pattern.test(source)) throw new Error("Registrazione sw.js versionata non trovata in index.html");
  return String(source).replace(
    pattern,
    `navigator.serviceWorker.register("./sw.js?v=${version}"`
  );
}

function stampFiles({ version = resolveBuildVersion(), swPath = SW_PATH, indexPath = INDEX_PATH } = {}) {
  const swBefore = fs.readFileSync(swPath, "utf8");
  const indexBefore = fs.readFileSync(indexPath, "utf8");
  const swAfter = stampServiceWorker(swBefore, version);
  const indexAfter = stampIndex(indexBefore, version);
  fs.writeFileSync(swPath, swAfter, "utf8");
  fs.writeFileSync(indexPath, indexAfter, "utf8");
  return { version, swChanged: swAfter !== swBefore, indexChanged: indexAfter !== indexBefore };
}

if (require.main === module) {
  const result = stampFiles();
  console.log(`PWA build version: ${result.version}`);
  console.log(`sw.js: ${result.swChanged ? "stampato" : "già aggiornato"}`);
  console.log(`index.html: ${result.indexChanged ? "stampato" : "già aggiornato"}`);
}

module.exports = {
  normalizeBuildVersion,
  resolveBuildVersion,
  stampServiceWorker,
  stampIndex,
  stampFiles
};

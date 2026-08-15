#!/usr/bin/env node
"use strict";

const fs = require("node:fs");

function replaceOnce(path, before, after) {
  const source = fs.readFileSync(path, "utf8");
  if (!source.includes(before)) throw new Error(`Blocco atteso non trovato in ${path}`);
  if (source.indexOf(before) !== source.lastIndexOf(before)) throw new Error(`Blocco ambiguo in ${path}`);
  fs.writeFileSync(path, source.replace(before, after));
}

replaceOnce(
  "auth-login-fix.js",
  `    const currentDoc = await currentRef.get();\n    if (currentDoc.exists) return;`,
  `    const currentDoc = await currentRef.get();\n    window.HeraPlatformProfileBootstrap = {\n      uid: user.uid,\n      exists: currentDoc.exists === true,\n      data: currentDoc.exists ? (currentDoc.data() || {}) : null,\n      loadedAt: Date.now()\n    };\n    if (currentDoc.exists) return;`
);

replaceOnce(
  "approval-access.js",
  `  async function ensureProfile(firebaseUser) {\n    const ref = db.collection("platformUsers").doc(firebaseUser.uid);\n    return db.runTransaction(async (transaction) => {`,
  `  async function ensureProfile(firebaseUser) {\n    const ref = db.collection("platformUsers").doc(firebaseUser.uid);\n    const bootstrap = window.HeraPlatformProfileBootstrap;\n    const bootstrapAgeMs = Date.now() - Number(bootstrap?.loadedAt || 0);\n    if (\n      bootstrap?.uid === firebaseUser.uid\n      && bootstrap.exists === true\n      && bootstrapAgeMs >= 0\n      && bootstrapAgeMs <= 2000\n    ) {\n      window.HeraPlatformProfileBootstrap = null;\n      return bootstrap.data && typeof bootstrap.data === "object" ? bootstrap.data : {};\n    }\n    return db.runTransaction(async (transaction) => {`
);

const checkPath = "scripts/check-platform-profile-bootstrap-reuse.js";
const checkSource = [
  '"use strict";',
  '',
  'const fs = require("node:fs");',
  'const auth = fs.readFileSync("auth-login-fix.js", "utf8");',
  'const approval = fs.readFileSync("approval-access.js", "utf8");',
  '',
  'const checks = [',
  '  [auth.includes("window.HeraPlatformProfileBootstrap = {"), "auth deve pubblicare il profilo appena letto"],',
  '  [auth.includes("exists: currentDoc.exists === true"), "la cache deve distinguere documenti esistenti"],',
  '  [approval.includes("bootstrap?.uid === firebaseUser.uid"), "approval deve riusare solo il profilo dello stesso UID"],',
  '  [approval.includes("bootstrap.exists === true"), "approval non deve riusare cache di profili mancanti"],',
  '  [approval.includes("bootstrapAgeMs <= 2000"), "la cache bootstrap deve scadere rapidamente"],',
  '  [approval.includes("window.HeraPlatformProfileBootstrap = null"), "la cache bootstrap deve essere one-shot"],',
  '  [approval.includes("return db.runTransaction(async (transaction) =>"), "il fallback transaction Firestore deve restare intatto"],',
  '  [approval.includes("const snapshot = await transaction.get(ref)"), "la lettura transaction originale deve restare disponibile"],',
  '];',
  '',
  'const failed = checks.filter(([ok]) => !ok).map(([, message]) => message);',
  'if (failed.length) {',
  '  console.error("Controllo riuso profilo bootstrap NON superato:");',
  '  failed.forEach((message) => console.error("- " + message));',
  '  process.exit(1);',
  '}',
  'console.log("Controllo riuso profilo bootstrap superato.");',
  ''
].join("\n");
fs.writeFileSync(checkPath, checkSource);

const packagePath = "package.json";
const pkg = JSON.parse(fs.readFileSync(packagePath, "utf8"));
pkg.scripts = pkg.scripts || {};
pkg.scripts["check:platform-profile-bootstrap"] = "node scripts/check-platform-profile-bootstrap-reuse.js";
if (!String(pkg.scripts["check:syntax"] || "").includes("check:platform-profile-bootstrap")) {
  pkg.scripts["check:syntax"] = `${pkg.scripts["check:syntax"]} && npm run check:platform-profile-bootstrap`;
}
fs.writeFileSync(packagePath, JSON.stringify(pkg, null, 2) + "\n");

console.log("Riuso one-shot del profilo platformUsers configurato.");

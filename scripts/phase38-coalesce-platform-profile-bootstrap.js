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
fs.writeFileSync(checkPath, `"use strict";\n\nconst fs = require("node:fs");\nconst auth = fs.readFileSync("auth-login-fix.js", "utf8");\nconst approval = fs.readFileSync("approval-access.js", "utf8");\n\nconst checks = [\n  [auth.includes("window.HeraPlatformProfileBootstrap = {"), "auth deve pubblicare il profilo appena letto"],\n  [auth.includes("exists: currentDoc.exists === true"), "la cache deve distinguere documenti esistenti"],\n  [approval.includes("bootstrap?.uid === firebaseUser.uid"), "approval deve riusare solo il profilo dello stesso UID"],\n  [approval.includes("bootstrap.exists === true"), "approval non deve riusare cache di profili mancanti"],\n  [approval.includes("bootstrapAgeMs <= 2000"), "la cache bootstrap deve scadere rapidamente"],\n  [approval.includes("window.HeraPlatformProfileBootstrap = null"), "la cache bootstrap deve essere one-shot"],\n  [approval.includes("return db.runTransaction(async (transaction) =>"), "il fallback transaction Firestore deve restare intatto"],\n  [approval.includes("const snapshot = await transaction.get(ref)"), "la lettura transaction originale deve restare disponibile"],\n  [!auth.includes("setInterval("), "l'ottimizzazione login non deve introdurre polling"],\n];\n\nconst failed = checks.filter(([ok]) => !ok).map(([, message]) => message);\nif (failed.length) {\n  console.error("Controllo riuso profilo bootstrap NON superato:");\n  failed.forEach((message) => console.error(`- ${message}`));\n  process.exit(1);\n}\nconsole.log("Controllo riuso profilo bootstrap superato.");\n`);

const packagePath = "package.json";
const pkg = JSON.parse(fs.readFileSync(packagePath, "utf8"));
pkg.scripts = pkg.scripts || {};
pkg.scripts["check:platform-profile-bootstrap"] = "node scripts/check-platform-profile-bootstrap-reuse.js";
if (!String(pkg.scripts["check:syntax"] || "").includes("check:platform-profile-bootstrap")) {
  pkg.scripts["check:syntax"] = `${pkg.scripts["check:syntax"]} && npm run check:platform-profile-bootstrap`;
}
fs.writeFileSync(packagePath, JSON.stringify(pkg, null, 2) + "\n");

console.log("Riuso one-shot del profilo platformUsers configurato.");

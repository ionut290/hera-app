"use strict";

const fs = require("node:fs");
const auth = fs.readFileSync("auth-login-fix.js", "utf8");
const approval = fs.readFileSync("approval-access.js", "utf8");

const checks = [
  [auth.includes("window.HeraPlatformProfileBootstrap = {"), "auth deve pubblicare il profilo appena letto"],
  [auth.includes("exists: currentDoc.exists === true"), "la cache deve distinguere documenti esistenti"],
  [approval.includes("bootstrap?.uid === firebaseUser.uid"), "approval deve riusare solo il profilo dello stesso UID"],
  [approval.includes("bootstrap.exists === true"), "approval non deve riusare cache di profili mancanti"],
  [approval.includes("bootstrapAgeMs <= 2000"), "la cache bootstrap deve scadere rapidamente"],
  [approval.includes("window.HeraPlatformProfileBootstrap = null"), "la cache bootstrap deve essere one-shot"],
  [approval.includes("return db.runTransaction(async (transaction) =>"), "il fallback transaction Firestore deve restare intatto"],
  [approval.includes("const snapshot = await transaction.get(ref)"), "la lettura transaction originale deve restare disponibile"],
];

const failed = checks.filter(([ok]) => !ok).map(([, message]) => message);
if (failed.length) {
  console.error("Controllo riuso profilo bootstrap NON superato:");
  failed.forEach((message) => console.error("- " + message));
  process.exit(1);
}
console.log("Controllo riuso profilo bootstrap superato.");

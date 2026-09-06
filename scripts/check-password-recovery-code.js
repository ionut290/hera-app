#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const core = require("../functions/password-recovery-code-core");

const read = (path) => fs.readFileSync(path, "utf8");
const backend = read("functions/password-recovery-code.js");
const main = read("functions/main.js");
const client = read("password-recovery-code.js");
const login = read("login-retry-fix.js");
const html = read("index.html");
const workflow = read(".github/workflows/deploy-firebase-functions.yml");
const capacitor = read("scripts/prepare-capacitor-web.js");
const serviceWorker = read("sw.js");

assert.equal(core.isRecoveryCodeCandidate("REC-corto"), false);
assert.equal(core.isRecoveryCodeCandidate("REC-Giardino-8427!"), true);
assert.equal(core.isRecoveryCodeCandidate("REC-Giardino con spazi"), false);
assert.equal(core.isValidNewPassword("123456789"), false);
assert.equal(core.isValidNewPassword("NuovaPass10!"), true);
const salt = core.createSalt();
const hash = core.hashRecoveryCode("REC-Giardino-8427!", salt);
assert.notEqual(hash, "REC-Giardino-8427!");
assert.equal(core.secureEqual(hash, core.hashRecoveryCode("REC-Giardino-8427!", salt)), true);
assert.equal(core.secureEqual(hash, core.hashRecoveryCode("REC-Altro-Codice1!", salt)), false);

for (const expected of [
  "setPasswordRecoveryCode",
  "getPasswordRecoveryCodeStatus",
  "startPasswordRecoveryWithCode",
  "completePasswordRecoveryWithCode",
  "hashRecoveryCode",
  "secureEqual",
  "MAX_FAILURES = 5",
  "LOCK_MS = 30 * 60 * 1000",
  "CHALLENGE_MS = 10 * 60 * 1000",
  "profileIsAdministrator",
  "emailIsConfiguredAdministrator",
  "profileIsActive",
  "admin.auth().updateUser",
  "admin.auth().revokeRefreshTokens",
  "password-recovered-with-shared-code"
]) assert.ok(backend.includes(expected), `Protezione backend mancante: ${expected}`);

assert.ok(main.includes('require("./password-recovery-code")'));
assert.ok(!backend.includes("deleteUser("), "Il recupero non deve eliminare account.");
assert.ok(!backend.includes("codePlain"), "Il codice non deve essere salvato in chiaro.");

for (const expected of [
  "REC-",
  "handleLoginCode",
  'callable("startPasswordRecoveryWithCode")',
  'callable("completePasswordRecoveryWithCode")',
  'callable("setPasswordRecoveryCode")',
  "HeraLoginCredentialVault?.capturePendingCredential"
]) assert.ok(client.includes(expected), `Flusso client mancante: ${expected}`);
assert.ok(client.includes('id="password-recovery-code-accept" type="checkbox" required'),
  "Manca la conferma esplicita prima della sostituzione della password.");
assert.ok(client.includes('adminList.insertAdjacentElement("afterend", section)'),
  "Il codice unico deve comparire subito sotto gli amministratori autorizzati.");

const recoveryIndex = login.indexOf("VargaPasswordRecovery?.isRecoveryCodeCandidate");
const vaultIndex = login.indexOf("HeraLoginCredentialVault?.capturePendingCredential");
const firebaseLoginIndex = login.indexOf("auth.signInWithEmailAndPassword", recoveryIndex);
assert.ok(recoveryIndex >= 0 && recoveryIndex < vaultIndex && recoveryIndex < firebaseLoginIndex,
  "Il codice deve essere riconosciuto prima del salvataggio credenziali e del login Firebase.");
assert.ok(html.indexOf("password-recovery-code.js") < html.indexOf("login-retry-fix.js"));
assert.ok(capacitor.includes('"password-recovery-code.js"'));
assert.ok(serviceWorker.includes("password-recovery-code.js"));
for (const functionName of [
  "setPasswordRecoveryCode",
  "getPasswordRecoveryCodeStatus",
  "startPasswordRecoveryWithCode",
  "completePasswordRecoveryWithCode"
]) assert.ok(workflow.includes(`functions:${functionName}`) && workflow.includes(`"${functionName}"`), `Deploy mancante: ${functionName}`);

console.log("Secure shared password recovery code checks passed.");

#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const vaultSource = fs.readFileSync("login-password-vault.js", "utf8");
const retrySource = fs.readFileSync("login-retry-fix.js", "utf8");
const appSource = fs.readFileSync("app.js", "utf8");
const nativeVaultSource = fs.readFileSync(
  "android/app/src/main/java/it/vargacantieri/hera/biometric/HeraCredentialVaultPlugin.java",
  "utf8"
);

assert.match(vaultSource, /capturePendingCredential,\s*\n\s*migrateLegacyAndroidAccounts/);
assert.match(
  retrySource,
  /HeraLoginCredentialVault\?\.capturePendingCredential\?\.\(\{ email, password \}\)[\s\S]*?signInWithEmailAndPassword\(email, password\)/,
  "La credenziale deve essere acquisita prima del tentativo Firebase."
);

const logoutStart = appSource.indexOf("async function logout()");
const logoutEnd = appSource.indexOf("\n}\n", logoutStart) + 2;
assert.ok(logoutStart >= 0 && logoutEnd > logoutStart, "Funzione logout non trovata.");
const logoutSource = appSource.slice(logoutStart, logoutEnd);
assert.doesNotMatch(
  logoutSource,
  /heraSavedLoginAccountsV1|HeraCredentialVault|deleteCredential|localStorage\.clear|sessionStorage\.clear/,
  "Il logout non deve cancellare le credenziali salvate."
);

assert.match(nativeVaultSource, /AES\/GCM\/NoPadding/);
assert.match(nativeVaultSource, /AndroidKeyStore/);
assert.match(nativeVaultSource, /getSharedPreferences\(PREFS, Context\.MODE_PRIVATE\)/);
assert.match(nativeVaultSource, /BIOMETRIC_STRONG/);

const storage = new Map();
const elements = {
  "saved-password-remember": { checked: true }
};
let domReady = null;
let authStateChanged = null;

const localStorage = {
  getItem: (key) => storage.has(key) ? storage.get(key) : null,
  setItem: (key, value) => storage.set(key, String(value)),
  removeItem: (key) => storage.delete(key)
};
const document = {
  readyState: "loading",
  body: {},
  getElementById: (id) => elements[id] || null,
  addEventListener: (type, callback) => {
    if (type === "DOMContentLoaded") domReady = callback;
  }
};
const auth = {
  onAuthStateChanged: (callback) => {
    authStateChanged = callback;
    return () => {};
  }
};
const window = {
  firebase: { auth: () => auth },
  Capacitor: null,
  confirm: () => true,
  setInterval: () => 1,
  clearInterval: () => {}
};

const context = vm.createContext({
  window,
  document,
  localStorage,
  firebase: window.firebase,
  MutationObserver: class { observe() {} },
  console,
  Date,
  Object,
  JSON,
  Array,
  String,
  Boolean,
  Promise
});

vm.runInContext(vaultSource, context, { filename: "login-password-vault.js" });
assert.equal(typeof domReady, "function", "Installazione vault non pianificata.");
domReady();
assert.equal(typeof authStateChanged, "function", "Hook Firebase del vault non installato.");

assert.equal(
  window.HeraLoginCredentialVault.capturePendingCredential({
    email: " Operatore@Example.it ",
    password: "Password-di-prova"
  }),
  true
);
authStateChanged({ email: "operatore@example.it" });

setImmediate(() => {
  const saved = JSON.parse(localStorage.getItem("heraSavedLoginAccountsV1") || "[]");
  assert.equal(saved.length, 1, "La credenziale non e stata salvata dopo il login riuscito.");
  assert.equal(saved[0].email, "operatore@example.it");
  assert.equal(saved[0].password, "Password-di-prova");

  localStorage.setItem("heraPersistedSessionV1", "sessione");
  localStorage.removeItem("heraPersistedSessionV1");
  assert.equal(
    JSON.parse(localStorage.getItem("heraSavedLoginAccountsV1") || "[]").length,
    1,
    "La credenziale deve restare disponibile dopo il logout."
  );

  console.log("Credential vault check passed: successful login saves credentials and logout preserves them.");
});

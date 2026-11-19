"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { execFileSync } = require("node:child_process");

const root = path.resolve(__dirname, "..");
function read(relative) {
  return fs.readFileSync(path.join(root, relative), "utf8");
}
function contains(relative, fragment) {
  assert.ok(read(relative).includes(fragment), relative + " missing: " + fragment);
}

execFileSync(process.execPath, ["--check", path.join(root, "account-privacy-controls.js")]);
execFileSync(process.execPath, ["--check", path.join(root, "functions/account-deletion.js")]);

contains("index.html", 'src="account-privacy-controls.js?');
contains("account-privacy-controls.js", '"auth-apple-login-btn"');
contains("account-privacy-controls.js", 'nativePlatform() === "ios"');
contains("account-privacy-controls.js", "signInWithApple({ skipNativeAuth: true })");
contains("account-privacy-controls.js", "new firebase.auth.OAuthProvider");
contains("account-privacy-controls.js", "reauthenticateWithCredential");
contains("account-privacy-controls.js", '"varga-account-delete-menu"');
contains("account-privacy-controls.js", '"varga-account-delete-pending"');
contains("account-privacy-controls.js", 'httpsCallable("deleteOwnVargaAccount")');
contains("functions/main.js", 'require("./account-deletion")');
contains("functions/account-deletion.js", "request.auth");
contains("functions/account-deletion.js", "auth.token?.auth_time");
contains("functions/account-deletion.js", "await admin.auth().deleteUser(uid)");
contains("functions/account-deletion.js", 'visibility === "personal"');
contains("functions/account-deletion.js", 'path.startsWith(prefix)');
contains("functions/account-deletion.js", "Business records shared with a company are not deleted");

const cfg = JSON.parse(read("capacitor.config.json"));
assert.equal(cfg.appId, "it.vargacantieri.hera");
assert.equal(cfg.plugins.FirebaseAuthentication.skipNativeAuth, true);
assert.ok(cfg.plugins.FirebaseAuthentication.providers.includes("apple.com"));
console.log("PASS: App Review account deletion and Apple Sign In source checks.");
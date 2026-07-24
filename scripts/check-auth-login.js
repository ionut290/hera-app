const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const indexSource = fs.readFileSync(path.join(root, "index.html"), "utf8");
const fixSource = fs.readFileSync(path.join(root, "auth-login-fix.js"), "utf8");
const workflowSource = fs.readFileSync(
  path.join(root, ".github", "workflows", "build-android-aab.yml"),
  "utf8"
);

if (!indexSource.includes('<script src="auth-login-fix.js?v=20260724a"></script>')) {
  throw new Error("auth-login-fix.js non viene caricato da index.html.");
}

if (!fixSource.includes("signInWithPopup(provider)")) {
  throw new Error("La correzione login non usa signInWithPopup.");
}

if (/signInWithRedirect\s*\(/.test(fixSource)) {
  throw new Error("La correzione login contiene signInWithRedirect.");
}

if (!fixSource.includes("stopImmediatePropagation")) {
  throw new Error("La correzione non intercetta il vecchio gestore login.");
}

if (!workflowSource.includes("auth-login-fix.js")) {
  throw new Error("Il workflow Android non include auth-login-fix.js.");
}

console.log("Google login check passed: direct popup handler included in web and Android builds.");

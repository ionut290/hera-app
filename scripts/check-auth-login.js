const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const indexSource = fs.readFileSync(path.join(root, "index.html"), "utf8");
const fixSource = fs.readFileSync(path.join(root, "auth-login-fix.js"), "utf8");
const workflowSource = fs.readFileSync(
  path.join(root, ".github", "workflows", "build-android-aab.yml"),
  "utf8"
);

if (!indexSource.includes('<script src="auth-login-fix.js?v=20260724b"></script>')) {
  throw new Error("auth-login-fix.js non viene caricato da index.html.");
}

if (!fixSource.includes('registerPlugin("FirebaseAuthentication")')) {
  throw new Error("Il plugin Firebase Authentication non viene registrato.");
}

if (!fixSource.includes("signInWithGoogle({ skipNativeAuth: true })")) {
  throw new Error("Il login Google nativo Android non viene avviato.");
}

if (!fixSource.includes("signInWithCredential(credential)")) {
  throw new Error("La credenziale Google nativa non viene trasferita a Firebase Web.");
}

if (!fixSource.includes("signInWithPopup(provider)")) {
  throw new Error("Il login web non mantiene signInWithPopup.");
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

if (!workflowSource.includes("ANDROID_GOOGLE_SERVICES_JSON_BASE64")) {
  throw new Error("Il workflow non ripristina google-services.json.");
}

if (!workflowSource.includes("rgcfaIncludeGoogle")) {
  throw new Error("Il workflow non abilita le dipendenze native Google.");
}

console.log("Google login check passed: native Android and web handlers are configured.");

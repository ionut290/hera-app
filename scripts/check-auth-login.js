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

if (!fixSource.includes("configureAndroidEmailPasswordOnly")) {
  throw new Error("La modalità Android email/password non viene configurata.");
}

if (!fixSource.includes('googleLoginButton.hidden = true')) {
  throw new Error("Il pulsante Google non viene nascosto su Android.");
}

if (!fixSource.includes('emailLoginButton.textContent = "Accedi"')) {
  throw new Error("Il pulsante email/password Android non ha il testo corretto.");
}

if (!fixSource.includes('message.textContent = "Accedi con la tua email e password."')) {
  throw new Error("Il messaggio login Android non richiede email e password.");
}

if (!fixSource.includes("Nell'app Android è disponibile solo l'accesso con email e password.")) {
  throw new Error("Il login Google non è bloccato esplicitamente su Android.");
}

if (!fixSource.includes("signInWithPopup(provider)")) {
  throw new Error("Il login web non mantiene signInWithPopup.");
}

if (/signInWithRedirect\s*\(/.test(fixSource)) {
  throw new Error("La correzione login contiene signInWithRedirect.");
}

if (!fixSource.includes("window.loginWithGoogle = function loginWithGoogleFixed()")) {
  throw new Error("La funzione centrale loginWithGoogle non viene sostituita.");
}

if (!fixSource.includes("stopImmediatePropagation")) {
  throw new Error("La correzione non intercetta il vecchio gestore login.");
}

if (!workflowSource.includes("auth-login-fix.js")) {
  throw new Error("Il workflow Android non include auth-login-fix.js.");
}

console.log("Login check passed: Android uses email/password only; web keeps Google login.");

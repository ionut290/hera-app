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

if (!fixSource.includes("installProfileAccessGuard")) {
  throw new Error("Il controllo profilo per accesso email/password non viene installato.");
}

if (!fixSource.includes("ensurePlatformProfileForAuthenticatedUser")) {
  throw new Error("Il profilo platformUsers non viene preparato prima dell'avvio app.");
}

if (!fixSource.includes('collection("platformUsers")')) {
  throw new Error("La correzione non usa la raccolta platformUsers.");
}

if (!fixSource.includes('where("email", "==", email)')) {
  throw new Error("La correzione non recupera il profilo esistente tramite email.");
}

if (!fixSource.includes("profileMigratedByEmail")) {
  throw new Error("La migrazione del profilo tramite email non viene registrata.");
}

if (!fixSource.includes("authInstance.onAuthStateChanged = function onAuthStateChangedWithProfile")) {
  throw new Error("app.js potrebbe controllare il profilo prima della sua preparazione.");
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

console.log("Login check passed: Android email/password inherits the same app profile and permissions.");
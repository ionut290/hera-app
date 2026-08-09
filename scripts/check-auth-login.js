const fs = require("fs");

const html = fs.readFileSync("index.html", "utf8");
const app = fs.readFileSync("app.js", "utf8");
const rules = fs.readFileSync("firestore.rules", "utf8");
const capacitor = fs.readFileSync("capacitor.config.json", "utf8");
const packageJson = fs.readFileSync("package.json", "utf8");
const androidBundle = fs.readFileSync("scripts/prepare-capacitor-web.js", "utf8");

for (const forbidden of [
  "Login con Google",
  "REGISTRATI CON GOOGLE",
  'id="auth-create-account-btn"',
  'id="auth-request-password-btn"',
  'src="auth-login-fix.js',
  'src="first-login-password.js',
  'src="login-retry-fix.js'
]) {
  if (html.includes(forbidden)) throw new Error(`Accesso legacy ancora presente in index.html: ${forbidden}`);
}

for (const expected of [
  'id="login-btn" class="btn btn-primary" type="button">Accesso amministratore</button>',
  'id="auth-email-input" type="text" autocomplete="username"',
  'id="auth-password-input" type="password"',
  'id="auth-admin-cancel-btn"',
  "Continua con accesso libero"
]) {
  if (!html.includes(expected)) throw new Error(`Interfaccia accesso libero/admin incompleta: ${expected}`);
}

for (const expected of [
  "await auth.signInAnonymously()",
  "function isPublicAccessUser",
  "function openAdminLogin",
  "await auth.signInWithEmailAndPassword(email, password)",
  "Username amministratore non riconosciuto",
  "Accesso libero attivo"
]) {
  if (!app.includes(expected)) throw new Error(`Logica accesso libero/admin incompleta: ${expected}`);
}

for (const forbidden of ["function loginWithGoogle", "function switchGoogleAccount", "ui.authGateLoginBtn?.addEventListener"]) {
  if (app.includes(forbidden)) throw new Error(`Login Google ancora attivo in app.js: ${forbidden}`);
}

if (!rules.includes('request.auth.token.firebase.sign_in_provider == "anonymous"')) {
  throw new Error("Le regole Firestore non riconoscono l'operatore pubblico anonimo.");
}
if (!rules.includes("isAdmin() || isPublicOperator()")) {
  throw new Error("L'accesso pubblico non è incluso nella funzione signedIn().");
}
if (capacitor.includes("google.com") || packageJson.includes("@capacitor-firebase/authentication")) {
  throw new Error("La configurazione Android contiene ancora il provider Google.");
}
for (const obsolete of ["auth-login-fix.js", "approval-access.js", "first-login-password.js", "login-retry-fix.js"]) {
  if (androidBundle.includes(`\"${obsolete}\"`)) throw new Error(`Asset login legacy ancora obbligatorio nel bundle Android: ${obsolete}`);
}

console.log("Access check passed: ingresso libero anonimo e credenziali riservate agli amministratori.");

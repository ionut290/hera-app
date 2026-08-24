const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const indexSource = fs.readFileSync(path.join(root, "index.html"), "utf8");
const fixSource = fs.readFileSync(path.join(root, "auth-login-fix.js"), "utf8");
const workflowSource = fs.readFileSync(
  path.join(root, ".github", "workflows", "build-android-aab.yml"),
  "utf8"
);
const capacitorBundleSource = fs.readFileSync(
  path.join(root, "scripts", "prepare-capacitor-web.js"),
  "utf8"
);
const capacitorConfigSource = fs.readFileSync(path.join(root, "capacitor.config.json"), "utf8");
const packageSource = fs.readFileSync(path.join(root, "package.json"), "utf8");

const rulesSource = fs.readFileSync(path.join(root, "firestore.rules"), "utf8");
const appSource = fs.readFileSync(path.join(root, "app.js"), "utf8");
const rulesDeployWorkflowSource = fs.readFileSync(
  path.join(root, ".github", "workflows", "deploy-self-registration.yml"),
  "utf8"
);

if (!/<script\s+src=["'](?:\.\/)?auth-login-fix\.js\?v=[^"']+["']><\/script>/.test(indexSource)) {
  throw new Error("auth-login-fix.js non viene caricato da index.html con cache-busting.");
}

if (!fixSource.includes("configurePlatformLoginOptions")) {
  throw new Error("Le opzioni login non vengono configurate in base alla piattaforma.");
}

if (!rulesSource.includes('.data.get("banned", false) == true')) {
  throw new Error("Le regole negano i dati quando il nuovo profilo non contiene banned.");
}

if (rulesSource.includes(".data.banned == true")) {
  throw new Error("Le regole leggono ancora direttamente il campo banned opzionale.");
}

if (!rulesDeployWorkflowSource.includes("--only firestore:rules")) {
  throw new Error("Il workflow non distribuisce le regole Firestore.");
}

for (const obsolete of ["TESTER_TEMP_PASSWORD", "functions:registerTester"]) {
  if (rulesDeployWorkflowSource.includes(obsolete)) {
    throw new Error(`Il deploy regole dipende ancora dalla registrazione legacy: ${obsolete}`);
  }
}

if (!fixSource.includes('googleLoginButton.hidden = false')) {
  throw new Error("Il pulsante Google non resta disponibile come seconda scelta su Android.");
}

if (!fixSource.includes('divider.hidden = false')) {
  throw new Error("Il separatore della seconda scelta Google non resta visibile.");
}

if (!fixSource.includes('Accedi con Google oppure con la tua email e password.')) {
  throw new Error("Il login web non comunica chiaramente l’accesso Google.");
}

if (!fixSource.includes('emailLoginButton.textContent = nativeAndroid ? "Accedi" : "Entra"')) {
  throw new Error("Il pulsante email/password non cambia testo correttamente in base alla piattaforma.");
}

if (!fixSource.includes('"Accedi con email/username e password. In alternativa puoi usare Google."')) {
  throw new Error("Android non presenta email/username e password come accesso principale e Google come alternativa.");
}

for (const expected of [
  "signInWithNativeGoogle",
  "nativeAuth.signInWithGoogle({ skipNativeAuth: true })",
  "firebase.auth.GoogleAuthProvider.credential(idToken)",
  "firebase.auth().signInWithCredential(credential)",
  "firebase.auth.Auth.Persistence.LOCAL",
  "isNativeAndroid() ? signInWithNativeGoogle() : signInWithWebGoogle()"
]) {
  if (!fixSource.includes(expected)) throw new Error(`Login Google Android incompleto: ${expected}`);
}

if (fixSource.includes("Nell'app Android è disponibile solo l'accesso con email e password.")) {
  throw new Error("È ricomparso il blocco esplicito del login Google su Android.");
}

if (!packageSource.includes('"@capacitor-firebase/authentication"')) {
  throw new Error("Plugin Firebase Authentication Android non installato.");
}

if (!capacitorConfigSource.includes('"providers": ["google.com"]')) {
  throw new Error("Provider Google Android non configurato in Capacitor.");
}

if (!capacitorConfigSource.includes('"skipNativeAuth": true')) {
  throw new Error("La configurazione Capacitor non mantiene skipNativeAuth per il bridge con Firebase Web.");
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

if (!fixSource.includes("catch (lookupError)")) {
  throw new Error("La ricerca del profilo per email può ancora bloccare il primo accesso.");
}

for (const expected of [
  'role: "user"',
  'ruolo: "user"',
  "isAdmin: false",
  "admin: false"
]) {
  if (!fixSource.includes(expected)) {
    throw new Error(`Il profilo automatico non è un utente normale: ${expected}`);
  }
}

if (fixSource.includes("banned: false")) {
  throw new Error("Il profilo automatico contiene un campo vietato dalle regole Firestore.");
}

if (!fixSource.includes("authInstance.onAuthStateChanged = function onAuthStateChangedWithProfile")) {
  throw new Error("app.js potrebbe controllare il profilo prima della sua preparazione.");
}

for (const expected of [
  "isPersistedApprovalValid",
  "session.accessApproved !== false",
  "const hasSavedApproval = isPersistedApprovalValid(savedSession, user)",
  "if (!hasSavedApproval)",
  'currentAccountStatus !== "attivo"'
]) {
  if (!appSource.includes(expected)) {
    throw new Error(`L'accesso rapido per un utente gia autorizzato non e protetto correttamente: ${expected}`);
  }
}

for (const expected of [
  "requiresEmailVerification",
  "user.emailVerified === false",
  "effectiveUser = emailVerificationRequired ? null : user",
  "showEmailVerificationRequired",
  "Prima di accedere ai dati, apri l’email di verifica"
]) {
  if (!fixSource.includes(expected)) {
    throw new Error(`Il caricamento dati non attende la verifica email: ${expected}`);
  }
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

if (
  !workflowSource.includes("npm run android:aab:prepare")
  || !capacitorBundleSource.includes('"auth-login-fix.js"')
) {
  throw new Error("Il workflow Android non include auth-login-fix.js.");
}

console.log("Login check passed: email/password resta principale; Google è una seconda scelta funzionale su Android e web.");

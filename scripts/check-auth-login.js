const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const indexSource = fs.readFileSync(path.join(root, "index.html"), "utf8");
const fixSource = fs.readFileSync(path.join(root, "auth-login-fix.js"), "utf8");
const operaPersistenceSource = fs.readFileSync(path.join(root, "opera-auth-persistence-fix.js"), "utf8");
const workflowSource = fs.readFileSync(
  path.join(root, ".github", "workflows", "build-android-aab.yml"),
  "utf8"
);
const capacitorBundleSource = fs.readFileSync(
  path.join(root, "scripts", "prepare-capacitor-web.js"),
  "utf8"
);

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
if (fixSource.includes("await ensurePlatformProfileForAuthenticatedUser(user)")) {
  throw new Error("Il listener Auth attende ancora Firestore prima di comunicare il login ad app.js.");
}
if (!fixSource.includes("Promise.resolve(result)") || !fixSource.includes("ensurePlatformProfileForAuthenticatedUser(user)")) {
  throw new Error("La preparazione profilo non bloccante dopo il login è assente.");
}
const authWrapperStart = fixSource.indexOf("const wrapCallback");
const authCallbackIndex = fixSource.indexOf("callback(effectiveUser)", authWrapperStart);
const profilePreparationIndex = fixSource.indexOf("ensurePlatformProfileForAuthenticatedUser(user)", authWrapperStart);
if (authWrapperStart < 0 || authCallbackIndex < 0 || profilePreparationIndex < authCallbackIndex) {
  throw new Error("La preparazione profilo parte ancora prima della comunicazione del login ad app.js.");
}

for (const expected of [
  "firebase.initializeApp(window.firebaseConfig)",
  "auth/unsupported-persistence-type",
  "persistence.SESSION",
  "persistence.NONE",
  "__heraCompatiblePersistenceInstalled"
]) {
  if (!operaPersistenceSource.includes(expected)) {
    throw new Error(`Fallback persistenza Opera incompleto: ${expected}`);
  }
}
const firebaseConfigIndex = indexSource.indexOf('src="firebase-config.js"');
const operaFixIndex = indexSource.indexOf('src="opera-auth-persistence-fix.js');
const authLoginIndex = indexSource.indexOf('src="auth-login-fix.js');
if (!(firebaseConfigIndex >= 0 && operaFixIndex > firebaseConfigIndex && authLoginIndex > operaFixIndex)) {
  throw new Error("Il fallback persistenza Opera deve essere caricato dopo Firebase config e prima dei gestori login.");
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

if (!fixSource.includes('googleLoginButton.hidden = !webGoogleEnabled')) {
  throw new Error("La visibilità del pulsante Google non dipende correttamente dalla piattaforma.");
}

if (!fixSource.includes('divider.hidden = !webGoogleEnabled')) {
  throw new Error("Il separatore Google non viene ripristinato sul web.");
}

if (!fixSource.includes('Accedi con Google oppure con la tua email e password.')) {
  throw new Error("Il login web non comunica chiaramente l’accesso Google.");
}

if (!fixSource.includes('emailLoginButton.textContent = nativeAndroid ? "Accedi" : "Entra"')) {
  throw new Error("Il pulsante email/password non cambia testo correttamente in base alla piattaforma.");
}

if (
  !fixSource.includes("message.textContent = nativeAndroid")
  || !fixSource.includes('"Accedi con la tua email e password."')
) {
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

console.log("Login check passed: web keeps Google; Android remains email/password-only.");

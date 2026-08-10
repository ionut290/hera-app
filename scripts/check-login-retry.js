const fs = require("fs");

const html = fs.readFileSync("index.html", "utf8");
const script = fs.readFileSync("login-retry-fix.js", "utf8");
const style = fs.readFileSync("login-retry-fix.css", "utf8");
const functionsSource = fs.readFileSync("functions/index.js", "utf8");
const androidWorkflow = fs.readFileSync(".github/workflows/build-android-aab.yml", "utf8");
const capacitorBundle = fs.readFileSync("scripts/prepare-capacitor-web.js", "utf8");
const deployWorkflow = fs.readFileSync(".github/workflows/deploy-register-tester.yml", "utf8");

if (!script.includes('/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/')) {
  throw new Error("La validazione email non accetta indirizzi standard.");
}
if (script.includes('/^[^\\\\s@]+@[^\\\\s@]+\\\\.[^\\\\s@]+$/')) {
  throw new Error("La validazione email contiene escape duplicati.");
}

for (const expected of [
  "Email o password non corretta.",
  "Email non ancora verificata.",
  "event.stopImmediatePropagation()",
  "window.__heraEmailLoginHandlerInstalled = true",
  "if (navigator.onLine === false)",
  "setCompatibleAuthPersistence",
  "auth/unsupported-persistence-type",
  "persistence.NONE",
  "loginButton.disabled = false",
  'loginButton.textContent = "Entra"',
  "window.scrollTo({ left: 0",
  "openRegistrationDialog",
  "startRegistrationFromLogin",
  'getElementById("auth-create-account-btn")',
  'getElementById("registration-dialog")',
  "firstName",
  "lastName",
  "MIN_REGISTRATION_PASSWORD_LENGTH = 10",
  "Creazione account non riuscita. Riprova tra poco."
]) {
  if (!script.includes(expected)) throw new Error(`Retry login incompleto: ${expected}`);
}
const onlineCheck = script.indexOf("if (navigator.onLine === false)");
const firstRememberedLogin = script.indexOf("await tryRememberedLogin(auth, email, feedback)", script.indexOf("async function handleLogin"));
if (onlineCheck < 0 || firstRememberedLogin < onlineCheck) {
  throw new Error("Una sessione salvata può ancora evitare la verifica password mentre il dispositivo è online.");
}

for (const expected of [
  "createUserWithEmailAndPassword(email, chosenPassword)",
  "createdUser.sendEmailVerification()",
  "createSelfRegisteredProfile(createdUser",
  'collection("platformUsers").doc(user.uid).set',
  "verificationRequired: true",
  "auth.signOut()",
  "role: \"user\"",
  "isAdmin: false",
  "selfRegistered: true"
]) {
  if (!script.includes(expected)) throw new Error(`Registrazione Firebase client mancante: ${expected}`);
}
if (script.includes('httpsCallable("registerTester")')) {
  throw new Error("Il client dipende ancora dalla Cloud Function privata registerTester.");
}

for (const expected of [
  'id="auth-create-account-btn"',
  ">CREA NUOVO ACCOUNT</button>",
  'id="registration-dialog"',
  'id="registration-first-name"',
  'id="registration-last-name"',
  'id="registration-password-confirm"',
  'minlength="10" autocomplete="new-password"',
  "ti invieremo un’email per verificare il nuovo account"
]) {
  if (!html.includes(expected)) throw new Error(`Registrazione HTML incompleta: ${expected}`);
}

const loginRetryScriptTag = html.match(/<script[^>]+src=["']login-retry-fix\.js\?v=([^"']+)["'][^>]*><\/script>/i);
if (!loginRetryScriptTag) {
  throw new Error("Registrazione HTML incompleta: login-retry-fix.js non è caricato con una versione esplicita.");
}
if (!String(loginRetryScriptTag[1] || "").trim()) {
  throw new Error("Registrazione HTML incompleta: versione login-retry-fix.js vuota.");
}

if (!style.includes("overflow-wrap: anywhere") || !style.includes("max-width: 100vw")) {
  throw new Error("Layout mobile login non protetto.");
}

for (const asset of ["login-retry-fix.js", "login-retry-fix.css"]) {
  if (!html.includes(asset)) throw new Error(`${asset} non caricato da index.html.`);
  if (!androidWorkflow.includes("npm run android:aab:prepare") || !capacitorBundle.includes(`"${asset}"`)) {
    throw new Error(`${asset} non incluso nell'AAB.`);
  }
}

if (!style.includes(".registration-dialog form")) {
  throw new Error("Stile finestra registrazione mancante.");
}

if (!deployWorkflow.includes("firebase deploy --only functions:registerTester")) {
  throw new Error("Deploy della funzione legacy non configurato.");
}
for (const forbidden of [
  "google-github-actions/setup-gcloud",
  "gcloud functions add-iam-policy-binding registerTester",
  "--member=allUsers",
  "roles/cloudfunctions.invoker"
]) {
  if (deployWorkflow.includes(forbidden)) {
    throw new Error(`Permesso pubblico legacy ancora presente nel workflow: ${forbidden}`);
  }
}

const registerTesterStart = functionsSource.indexOf("exports.registerTester");
const registerTesterEnd = functionsSource.indexOf("function chunkItems", registerTesterStart);
if (registerTesterStart < 0 || registerTesterEnd < 0) {
  throw new Error("Funzione registerTester non trovata.");
}
const registerTesterSource = functionsSource.slice(registerTesterStart, registerTesterEnd);
for (const expected of [
  "password.length < 10",
  "mustChangePassword: false",
  "selfRegistered: true",
  "Creazione account non riuscita. Riprova tra poco."
]) {
  if (!registerTesterSource.includes(expected)) {
    throw new Error(`Backend registrazione legacy incompleto: ${expected}`);
  }
}
if (registerTesterSource.includes('invoker: "public"')) {
  throw new Error("La funzione legacy registerTester non deve essere pubblica.");
}
if (registerTesterSource.includes("functions.config().tester")) {
  throw new Error("La registrazione dipende ancora dalla vecchia password temporanea condivisa.");
}

console.log("Login retry and Firebase client registration check passed.");

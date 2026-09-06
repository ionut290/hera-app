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
  "Se l’account è già registrato, usa PASSWORD DIMENTICATA?.",
  "Email non ancora verificata.",
  "event.stopImmediatePropagation()",
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

const handleLoginStart = script.indexOf("  async function handleLogin(event)");
const handleLoginEnd = script.indexOf("  function initialize()", handleLoginStart);
if (handleLoginStart < 0 || handleLoginEnd < 0) {
  throw new Error("Gestore login email/password non trovato.");
}
const handleLoginSource = script.slice(handleLoginStart, handleLoginEnd);
if (handleLoginSource.includes("openRegistrationDialog")) {
  throw new Error("Un errore di accesso apre ancora automaticamente la registrazione.");
}
if (handleLoginSource.includes("Account non trovato")) {
  throw new Error("Il login distingue ancora in modo inaffidabile un account inesistente da una password errata.");
}
if (!handleLoginSource.includes("await auth.signInWithEmailAndPassword(email, password)")) {
  throw new Error("Accesso Firebase email/password mancante.");
}

for (const expected of [
  'document.body?.classList.add("auth-pending")',
  'document.body?.classList.remove("auth-required", "auth-banned")',
  'loader.classList.remove("hidden")',
  "loader.hidden = false"
]) {
  if (!script.includes(expected)) throw new Error(`Transizione post-login anti-schermata-bianca mancante: ${expected}`);
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
  "ti invieremo un’email per verificare il nuovo account",
  "login-retry-fix.js?v=20260906-recovery-code1"
]) {
  if (!html.includes(expected)) throw new Error(`Registrazione HTML incompleta: ${expected}`);
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

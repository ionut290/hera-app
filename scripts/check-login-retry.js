const fs = require("fs");

const html = fs.readFileSync("index.html", "utf8");
const script = fs.readFileSync("login-retry-fix.js", "utf8");
const style = fs.readFileSync("login-retry-fix.css", "utf8");
const functionsSource = fs.readFileSync("functions/index.js", "utf8");
const workflow = fs.readFileSync(".github/workflows/build-android-aab.yml", "utf8");

if (!script.includes('/^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$/')) {
  throw new Error("La validazione email non accetta indirizzi standard.");
}
if (script.includes('/^[^\\\\s@]+@[^\\\\s@]+\\\\.[^\\\\s@]+$/')) {
  throw new Error("La validazione email contiene escape duplicati.");
}

for (const expected of [
  "Email o password non corretta.",
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

for (const expected of [
  'functions("europe-west1").httpsCallable("registerTester")',
  "temporaryPassword: chosenPassword",
  "firstName",
  "lastName"
]) {
  if (!script.includes(expected)) throw new Error(`Auto-registrazione mancante: ${expected}`);
}

for (const expected of [
  'id="auth-create-account-btn"',
  '>CREA NUOVO ACCOUNT</button>',
  'id="registration-dialog"',
  'id="registration-first-name"',
  'id="registration-last-name"',
  'id="registration-password-confirm"',
  'minlength="10" autocomplete="new-password"',
  'login-retry-fix.js?v=20260726e'
]) {
  if (!html.includes(expected)) throw new Error(`Registrazione HTML incompleta: ${expected}`);
}

if (!style.includes("overflow-wrap: anywhere") || !style.includes("max-width: 100vw")) {
  throw new Error("Layout mobile login non protetto.");
}

for (const asset of ["login-retry-fix.js", "login-retry-fix.css"]) {
  if (!html.includes(asset)) throw new Error(`${asset} non caricato da index.html.`);
  if (!workflow.includes(asset)) throw new Error(`${asset} non incluso nell'AAB.`);
}

if (!style.includes(".registration-dialog form")) {
  throw new Error("Stile finestra registrazione mancante.");
}

for (const expected of [
  '.runWith({ invoker: "public" })',
  "password.length < 10",
  "mustChangePassword: false",
  "selfRegistered: true",
  "Creazione account non riuscita. Riprova tra poco."
]) {
  if (!functionsSource.includes(expected)) {
    throw new Error(`Backend registrazione incompleto: ${expected}`);
  }
}

const registerTesterStart = functionsSource.indexOf("exports.registerTester");
const registerTesterEnd = functionsSource.indexOf("function chunkItems", registerTesterStart);
if (registerTesterStart < 0 || registerTesterEnd < 0) {
  throw new Error("Funzione registerTester non trovata.");
}
const registerTesterSource = functionsSource.slice(registerTesterStart, registerTesterEnd);
if (registerTesterSource.includes("functions.config().tester")) {
  throw new Error("La registrazione dipende ancora dalla vecchia password temporanea condivisa.");
}

console.log("Login retry and new-user registration check passed.");

const fs = require("fs");

const html = fs.readFileSync("index.html", "utf8");
const client = fs.readFileSync("first-login-password.js", "utf8");
const functionsSource = fs.readFileSync("functions/index.js", "utf8");

for (const expected of [
  'id="auth-email-login-btn" class="btn btn-primary" type="submit">Entra</button>',
  'id="auth-request-password-btn"',
  '>PASSWORD DIMENTICATA?</button>',
  'id="first-password-dialog"',
  'src="first-login-password.js?v=20260726b"'
]) {
  if (!html.includes(expected)) throw new Error(`Elemento mancante in index.html: ${expected}`);
}

for (const expected of [
  "mustChangePassword",
  "updatePassword(nextPassword)",
  "passwordChangedAt",
  "sendPasswordResetEmail(email)",
  "sendPasswordResetInstructions",
  'auth.languageCode = "it"',
  "getPasswordResetContinueUrl",
  "handleCodeInApp: false",
  "showPasswordResetReturnNotice",
  "Password aggiornata. Inserisci email e nuova password",
  "passwordResetPending",
  "auth/network-request-failed",
  "auth/too-many-requests",
  "sendEmailVerification()",
  'dialog.addEventListener("cancel", (event) => event.preventDefault())'
]) {
  if (!client.includes(expected)) throw new Error(`Controllo client mancante: ${expected}`);
}

for (const expected of [
  "exports.createTesterAccounts",
  "generateTemporaryPassword",
  "mustChangePassword: true",
  "admin.auth().createUser",
  "admin.auth().updateUser"
]) {
  if (!functionsSource.includes(expected)) throw new Error(`Provisioning tester mancante: ${expected}`);
}

console.log("First-login password check passed.");

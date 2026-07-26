const fs = require("fs");

const html = fs.readFileSync("index.html", "utf8");
const script = fs.readFileSync("login-retry-fix.js", "utf8");
const style = fs.readFileSync("login-retry-fix.css", "utf8");
const workflow = fs.readFileSync(".github/workflows/build-android-aab.yml", "utf8");

for (const expected of [
  "Email o password non corretta.",
  "event.stopImmediatePropagation()",
  "loginButton.disabled = false",
  'loginButton.textContent = "Entra"',
  "window.scrollTo({ left: 0"
]) {
  if (!script.includes(expected)) throw new Error(`Retry login incompleto: ${expected}`);
}

for (const expected of [
  'functions("europe-west1").httpsCallable("registerTester")',
  "temporaryPassword: password"
]) {
  if (!script.includes(expected)) throw new Error(`Auto-registrazione mancante: ${expected}`);
}

if (!style.includes("overflow-wrap: anywhere") || !style.includes("max-width: 100vw")) {
  throw new Error("Layout mobile login non protetto.");
}

for (const asset of ["login-retry-fix.js", "login-retry-fix.css"]) {
  if (!html.includes(asset)) throw new Error(`${asset} non caricato da index.html.`);
  if (!workflow.includes(asset)) throw new Error(`${asset} non incluso nell'AAB.`);
}

console.log("Login retry check passed.");

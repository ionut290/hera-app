#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");

const read = (path) => fs.readFileSync(path, "utf8").replace(/\r\n/g, "\n");
const sha256 = (value) => crypto.createHash("sha256").update(value).digest("hex");
const login = read("login-retry-fix.js");
const contract = read("AUTH_STABILITY_CONTRACT.md");
const agents = read("AGENTS.md");

const blockStart = login.indexOf("  const LOGIN_STABILITY_CONTRACT");
const blockEnd = login.indexOf("  function friendlyLoginError", blockStart);
assert.ok(blockStart >= 0 && blockEnd > blockStart, "Blocco runtime stabilità login non trovato.");

const stabilityBlock = login.slice(blockStart, blockEnd);
assert.equal(
  sha256(stabilityBlock),
  "931287ba62a3f8beb4658c5545823a1964304def89529018677c21c103590728",
  "Il runtime di stabilità login è cambiato: serve il consenso esplicito del proprietario e l'aggiornamento consapevole del contratto."
);
assert.equal(
  sha256(contract),
  "53952299bbec367cee444ed753e9b49ce0f14661caba8be1d5fcf29426a0f46d",
  "Il contratto di stabilità login è cambiato senza aggiornare la sua impronta autorizzata."
);

for (const required of [
  "LOGIN_STABILITY_CONTRACT_V1",
  "hasVisibleAuthenticationSurface",
  "recoverVisibleAuthenticationSurface",
  "schedulePostLoginVisibilityChecks",
  "app-startup-loading",
  "access-approval-screen",
  "home-page"
]) {
  assert.ok(login.includes(required), `Protezione login incompleta: ${required}`);
}

assert.match(agents, /PROTEZIONE FORTE E MODIFICABILE DEL LOGIN/);
assert.match(agents, /consenso\s+esplicito/i);
assert.match(contract, /forte ma non irrevocabile/i);
assert.doesNotMatch(stabilityBlock, /firestore|collection\(|\.doc\(|onAuthStateChanged|onSnapshot|\.set\(|\.update\(/i,
  "La protezione visiva non deve aggiungere accessi o listener Firebase.");

function createClassList(initial = []) {
  const values = new Set(initial);
  return {
    add: (...names) => names.forEach((name) => values.add(name)),
    remove: (...names) => names.forEach((name) => values.delete(name)),
    contains: (name) => values.has(name)
  };
}

const elements = {
  "app-startup-loading": { hidden: true, classList: createClassList(["hidden"]), removeAttribute() {} },
  "auth-gate": { hidden: true, classList: createClassList(["hidden"]) },
  "access-approval-screen": { hidden: true, classList: createClassList(["hidden"]) },
  "home-page": { hidden: false, classList: createClassList() }
};
const body = { classList: createClassList(["auth-required"]) };
const documentHarness = { body, getElementById: (id) => elements[id] || null };
const scheduled = [];
const windowHarness = { setTimeout: (callback, delay) => scheduled.push({ callback, delay }) };
const warnings = [];
const runtime = Function("document", "window", "console", `${stabilityBlock}\nreturn { hasVisibleAuthenticationSurface, recoverVisibleAuthenticationSurface, schedulePostLoginVisibilityChecks };`)(
  documentHarness,
  windowHarness,
  { warn: (message) => warnings.push(message) }
);

assert.equal(runtime.hasVisibleAuthenticationSurface(), false, "Il test deve partire da una schermata bianca simulata.");
assert.equal(runtime.recoverVisibleAuthenticationSurface(), false, "Il recupero deve segnalare di avere corretto lo stato.");
assert.equal(elements["app-startup-loading"].hidden, false, "Il recupero deve mostrare lo splash.");
assert.equal(elements["app-startup-loading"].classList.contains("hidden"), false, "Lo splash non deve restare nascosto.");
assert.equal(body.classList.contains("auth-pending"), true, "La Home deve restare protetta durante il recupero.");
assert.equal(warnings.length, 1, "Il recupero deve lasciare una traccia diagnostica.");
runtime.schedulePostLoginVisibilityChecks();
assert.deepEqual(scheduled.map(({ delay }) => delay), [0, 1500, 4000, 8000], "Watchdog post-login incompleto.");

console.log("Login stability contract active: visible-state recovery protected and owner-authorized changes allowed.");

"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(path.join(__dirname, "../app.js"), "utf8");
const logger = source.slice(source.indexOf("async function logActivity("), source.indexOf("function filterActiveUsersLogs("));
const listener = source.slice(source.lastIndexOf('document.addEventListener("click", (event) => {'));
const writes = [];
let onClick;
const context = vm.createContext({
  db: { collection: (name) => ({ add: async (data) => { writes.push({ name, data }); } }) },
  currentUser: { uid: "test-operator", email: "test@example.invalid" },
  selectedCommessaId: "test-commessa",
  selectedCommessaName: "Test",
  canManageData: () => false,
  getCurrentViewName: () => "impianti",
  navigator: { userAgent: "test" },
  firebase: { firestore: { FieldValue: { serverTimestamp: () => "timestamp" } } },
  document: { addEventListener: (type, callback) => { assert.equal(type, "click"); onClick = callback; } },
  console
});
vm.runInContext(logger + listener, context);

(async () => {
  await context.logActivity("pressione_naviga", "Pressione NAVIGA");
  await context.logActivity(" pressione_naviga ", "Pressione NAVIGA");
  assert.equal(writes.length, 0, "Le chiamate dirette NAVIGA non devono scrivere dati");
  for (const label of ["Naviga", "NAVIGA VERSO L'IMPIANTO", "Confermo, procedi navigazione"]) {
    onClick({ target: { closest: () => ({ textContent: label, getAttribute: () => null }) } });
  }
  assert.equal(writes.length, 0, "I clic di navigazione non devono scrivere dati");
  await context.logActivity("errore_firestore", "Errore test");
  assert.equal(writes.length, 1, "Gli altri eventi devono mantenere il comportamento esistente");
  assert.equal(writes[0].data.actionType, "errore_firestore");
  console.log("✅ NAVIGA: nessuna scrittura dal clic o da chiamate dirette; altri eventi preservati.");
})().catch((error) => { console.error(error); process.exitCode = 1; });

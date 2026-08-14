#!/usr/bin/env node
const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("fatto-button-immediate.js", "utf8");

assert.match(source, /async function loadInreteWorkContext\(impianto\)/, "Manca lettura lavorazioni INRETE");
assert.match(source, /collection\("lavorazioni"\)\.where\("impiantoId", "==", plantId\)/, "Le lavorazioni devono essere filtrate per impianto");
assert.match(source, /async function savePartialWorkItem\(context, workItemId, impianto\)/, "Manca salvataggio singola lavorazione");
assert.match(source, /stato:\s*"FATTO"/, "La lavorazione selezionata deve diventare FATTO");
assert.match(source, /"PARZIALMENTE FATTO"/, "Manca stato parziale impianto");
assert.match(source, /numeroLavorazioniDaFare/, "Manca conteggio lavorazioni residue");
assert.match(source, /TUTTE LE LAVORAZIONI SONO FATTE/, "Manca scelta completamento totale");
assert.match(source, /ORDINARIO/, "Manca distinzione ordinario");
assert.match(source, /STRAORDINARIO/, "Manca distinzione straordinario");
assert.match(source, /const partial = await maybeHandlePartialInreteDone\(impianto\)/, "Il wrapper FATTO deve verificare prima il completamento parziale");
assert.ok(
  source.indexOf("const partial = await maybeHandlePartialInreteDone(impianto)") < source.indexOf("applyPermanentYellowFeedback(pressedButton, doneAt)"),
  "La scelta parziale deve avvenire prima del blocco definitivo del pulsante"
);
assert.match(source, /if \(saved\.allDone\) return \{ handled: false \}/, "L'ultima lavorazione deve rientrare nel FATTO completo");
assert.match(source, /🟡 LAVORAZIONE FATTA/, "Manca messaggio WhatsApp per FATTO parziale");

console.log("✅ Controlli FATTO parziale INRETE completati.");

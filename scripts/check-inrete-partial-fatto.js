#!/usr/bin/env node
const assert = require("node:assert/strict");
const fs = require("node:fs");

const source = fs.readFileSync("fatto-button-immediate.js", "utf8");

assert.match(source, /async function loadInreteWorkContext\(impianto\)/, "Manca lettura lavorazioni INRETE");
assert.match(source, /collection\("lavorazioni"\)\.where\("impiantoId", "==", plantId\)/, "Le lavorazioni devono essere filtrate per impianto");
assert.match(source, /PARZIALMENTE FATTO/, "Manca stato parziale impianto");
assert.match(source, /Cosa hai eseguito\?/, "Manca domanda di selezione intervento");
assert.match(source, /type = "checkbox"/, "La scelta deve usare caselle selezionabili");
assert.match(source, /Manutenzione ordinaria/, "Manca casella ordinaria");
assert.match(source, /Manutenzione straordinaria/, "Manca casella straordinaria");
assert.match(source, /Hai scelto solo la manutenzione ordinaria/, "Manca conferma ordinaria soltanto");
assert.match(source, /Hai scelto solo la manutenzione straordinaria/, "Manca conferma straordinaria soltanto");
assert.match(source, /TORNA INDIETRO/, "Manca azione torna indietro");
assert.match(source, /CONFERMA E INVIA/, "Manca conferma esplicita prima dell'invio");
assert.match(source, /Intervento eseguito: manutenzione ordinaria/, "Manca messaggio WhatsApp ordinaria");
assert.match(source, /Intervento eseguito: manutenzione straordinaria/, "Manca messaggio WhatsApp straordinaria");
assert.match(source, /Interventi eseguiti: manutenzione ordinaria e straordinaria/, "Manca messaggio combinato");
assert.match(source, /async function saveSelectedWorkItems\(context, selectedKinds, impianto\)/, "Manca salvataggio per tipo selezionato");
assert.match(source, /selectedKinds\.includes\(getWorkItemKind\(entry\)\)/, "Il salvataggio deve limitarsi alle lavorazioni selezionate");
assert.match(source, /numeroLavorazioniDaFare/, "Manca conteggio lavorazioni residue");
assert.match(source, /if \(saved\.allDone\) return \{ handled: false/, "L'ultima lavorazione deve rientrare nel FATTO completo");
assert.match(source, /const partial = await maybeHandlePartialInreteDone\(impianto\)/, "Il wrapper FATTO deve verificare prima il completamento parziale");
assert.ok(
  source.indexOf("const partial = await maybeHandlePartialInreteDone(impianto)") < source.indexOf("applyPermanentYellowFeedback(pressedButton, doneAt)"),
  "La scelta parziale deve avvenire prima del blocco definitivo del pulsante"
);
assert.doesNotMatch(source, /🟡 LAVORAZIONE FATTA/, "Il completamento parziale non deve usare il vecchio titolo generico");

console.log("✅ Controlli FATTO ordinario/straordinario completati.");

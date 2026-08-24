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

assert.match(source, /FALLBACK_STORE_KEY = "heraFattoSyncOperationsFallbackV1"/, "Manca il ripiego persistente della cassaforte FATTO");
assert.match(source, /async function enqueueSafely\(impianto, metadata = \{\}\)/, "La cassaforte deve avere un ingresso mai bloccante");
assert.match(source, /salvataggio locale non disponibile; il flusso operativo continua/, "Un errore della cassaforte non deve fermare FATTO");
assert.match(source, /const operation = await enqueueSafely\(impianto, \{[\s\S]*const result = await original\.call\(this, impianto, \.\.\.args\)/, "Il flusso FATTO deve continuare anche senza operazione in cassaforte");
assert.match(source, /if \(operation\) \{[\s\S]*result === true[\s\S]*setStatus\(operation, "FAILED"/, "Lo stato cassaforte deve essere aggiornato solo quando l'operazione esiste");
assert.match(source, /type: "IMPIANTO_FATTO_PARZIALE"/, "Il FATTO parziale fallito deve essere conservato con un tipo dedicato");
assert.match(source, /workItems: context\.items/, "La cassaforte parziale deve conservare le lavorazioni esatte");
assert.match(source, /salvataggio non riuscito; conservo nella cassaforte e apro Whazzup/, "Un errore parziale deve aprire Whazzup e attivare il recupero");
assert.match(source, /const opened = openPartialWhatsApp\(impianto, selectedKinds, selectedItems, doneAt, doneBy, remaining\)/, "Whazzup parziale deve aprirsi anche dopo errore di salvataggio");
assert.match(source, /operation\?\.type === "IMPIANTO_FATTO_PARZIALE"/, "Il recupero deve riconoscere le operazioni parziali");
assert.match(source, /return saveSelectedWorkItems\(context, operation\.selectedKinds \|\| \[\], operation\.impianto \|\| \{\}\)/, "Il recupero deve risalvare le lavorazioni selezionate senza riaprire Whazzup");

console.log("✅ Controlli FATTO ordinario/straordinario e cassaforte completati.");

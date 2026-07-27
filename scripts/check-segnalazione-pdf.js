"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const app = fs.readFileSync("app.js", "utf8");
const functionMatch = app.match(/async function generateSegnalazionePdf\(event\) \{([\s\S]*?)\n\}\n\nasync function shareSegnalazione/);

assert.ok(functionMatch, "generateSegnalazionePdf non trovata");
const implementation = functionMatch[1];

assert.match(implementation, /getSegnalazioneData\(\)/, "Il PDF deve usare i dati del modulo");
assert.match(implementation, /doc\.text\(/, "Il PDF deve contenere testo nativo");
assert.match(implementation, /doc\.rect\(/, "Il layout deve contenere rettangoli vettoriali");
assert.match(implementation, /doc\.line\(/, "Il layout deve contenere linee vettoriali");
assert.match(implementation, /splitTextToSize\(data\.descrizione/, "La descrizione deve andare a capo");
assert.match(implementation, /doc\.addPage\(/, "Le descrizioni lunghe devono creare altre pagine");
assert.match(implementation, /doc\.getNumberOfPages\(\)/, "Il PDF deve numerare tutte le pagine");
assert.match(implementation, /doc\.setProperties\(/, "Il PDF deve includere metadati");
assert.match(implementation, /doc\.output\("blob"\)/, "Il risultato deve essere un blob");
assert.match(implementation, /pdfBlob\.type !== "application\/pdf"/, "Il MIME type deve essere verificato");
assert.match(implementation, /signature !== "%PDF-"/, "La firma PDF deve essere verificata");
assert.match(implementation, /lastSegnalazionePdfBlob = pdfBlob/, "Il blob deve restare disponibile per la condivisione");
assert.match(implementation, /generateButton\.disabled = true/, "Il pulsante deve essere disabilitato durante la generazione");
assert.match(implementation, /finally \{/, "Il pulsante deve essere ripristinato anche in caso di errore");
assert.doesNotMatch(implementation, /html2canvas/, "La scheda non deve dipendere da html2canvas");
assert.doesNotMatch(implementation, /querySelector\("\.segnalazione-sheet"\)/, "Il modulo non deve essere catturato come immagine");
assert.doesNotMatch(implementation, /addImage\(imageData/, "L'intero modulo non deve essere inserito come singola immagine");

console.log("Segnalazione PDF checks passed.");

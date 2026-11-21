"use strict";

// Zero-network static contract for the public account-deletion page.
// Checks truthful live authenticated deletion, iOS Apple grant revocation,
// supported providers, no invented SLA and support contact limitations.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const html = fs.readFileSync(path.resolve(__dirname, "../delete-account.html"), "utf8");

assert.match(html, /Eliminazione account VARGA CANTIERI/);
assert.match(html, /La cancellazione autenticata è disponibile nell'app per gli utenti ordinari/);
assert.match(html, /Menu &gt; Personale e accessi &gt; Elimina il mio account/);
assert.match(html, /non sostituisce il percorso autenticato nell'app/);
assert.match(html, /recapito di assistenza/);
assert.match(html, /Google/);
assert.match(html, /Apple/);
assert.match(html, /password, codici di accesso/);
assert.match(html, /la revoca della credenziale viene eseguita nell'app iPhone prima della cancellazione/);
assert.match(html, /L'app mostra una conferma solo dopo la risposta positiva del server/);
assert.match(html, /Dati che potrebbero essere conservati/);
assert.doesNotMatch(html, /normalmente entro 30 giorni/i);
assert.doesNotMatch(html, /verificata e completata normalmente entro/i);
assert.doesNotMatch(html, /dall.indirizzo Google utilizzato per accedere/i);
assert.doesNotMatch(html, /Indica nome e cognome, email dell.account Google, azienda/i);
assert.doesNotMatch(html, /<a[^>]+>Invia richiesta<\/a>/i);
assert.doesNotMatch(html, /<form\b/i);
console.log("PASS: public page describes live authenticated deletion path, iPhone Apple revocation, support limitations and retention without invented SLA.");

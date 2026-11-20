"use strict";

// Zero-network static contract for the public account-deletion page.
// Checks honest status, both supported login providers, no invented SLA,
// and clear distinction between support contact and in-app deletion.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const html = fs.readFileSync(path.resolve(__dirname, "../delete-account.html"), "utf8");

assert.match(html, /Eliminazione account VARGA CANTIERI/);
assert.match(html, /non è ancora attiva nella versione distribuita/);
assert.match(html, /non sostituisce il percorso di cancellazione in-app/);
assert.match(html, /recapito di assistenza/);
assert.match(html, /Google/);
assert.match(html, /Apple/);
assert.match(html, /password, codici di accesso/);
assert.match(html, /Non è ancora stata convalidata una tempistica/);
assert.match(html, /completamento deve essere confermato espressamente/);
assert.match(html, /Dati che potrebbero essere conservati/);
assert.doesNotMatch(html, /normalmente entro 30 giorni/i);
assert.doesNotMatch(html, /verificata e completata normalmente entro/i);
assert.doesNotMatch(html, /dall.indirizzo Google utilizzato per accedere/i);
assert.doesNotMatch(html, /Indica nome e cognome, email dell.account Google, azienda/i);
assert.doesNotMatch(html, /<a[^>]+>Invia richiesta<\/a>/i);
assert.doesNotMatch(html, /<form\b/i);
console.log("PASS: public deletion page distinguishes Apple/Google, does not fake live in-app deletion, approved SLA or backend fulfillment.");

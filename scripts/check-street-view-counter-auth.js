'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const sourcePath = path.resolve(__dirname, '..', 'street-view-cards.js');
const source = fs.readFileSync(sourcePath, 'utf8');

new vm.Script(source, { filename: sourcePath });

assert.match(source, /function\s+waitForAuthenticatedUser\s*\(/, 'Deve attendere il ripristino della sessione Firebase');
assert.match(source, /await\s+waitForAuthenticatedUser\s*\(\s*\)/, 'Il contatore deve attendere la sessione prima della transazione');
assert.match(source, /authenticatedUser\.getIdToken\?\.\(true\)/, 'Un errore di autenticazione deve aggiornare il token una sola volta');
assert.match(source, /if\s*\(!isAuthenticationCounterError\(error\)\)\s*throw error/, 'Gli errori non legati all’autenticazione non devono creare tentativi duplicati');

const calls = source.match(/runSharedCounterTransaction\(firestore, ref, user, monthKey\)/g) || [];
assert.equal(calls.length, 3, 'Devono esistere soltanto la definizione, il tentativo normale e il singolo retry autenticato');
assert.doesNotMatch(source, /setInterval\([^)]*reserveSharedMonthlySlot/, 'Il contatore non deve usare polling');
assert.doesNotMatch(source, /onSnapshot\(/, 'Street View non deve aggiungere listener Firestore');

console.log('OK: Street View attende Firebase e ritenta il contatore soltanto dopo un errore di autenticazione.');

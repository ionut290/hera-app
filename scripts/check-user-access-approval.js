#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const client = fs.readFileSync("approval-access.js", "utf8");
const backend = fs.readFileSync("functions/user-access-approval.js", "utf8");
const entrypoint = fs.readFileSync("functions/main.js", "utf8");
const workflow = fs.readFileSync(".github/workflows/deploy-firebase-functions.yml", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");
const index = fs.readFileSync("index.html", "utf8");

assert.match(client, /httpsCallable\("approveUserAccess"\)/, "Sblocca utente usa il callable amministrativo");
assert.match(client, /SBLOCCO E INVIO EMAIL/, "Il pulsante mostra lo stato di avanzamento");
assert.match(client, /access-approval-result-whatsapp/, "Dopo lo sblocco è disponibile l’invio WhatsApp");
assert.match(client, /access-approval-result-email/, "Dopo lo sblocco è disponibile l’invio email manuale");
assert.match(client, /mailto:/, "L’email manuale contiene il messaggio completo senza servizi esterni obbligatori");
assert.match(client, /COPIA MESSAGGIO/, "Il messaggio può essere copiato come recupero manuale");
assert.match(client, /firebase\.app\(\)\.functions\(REGION\)/, "Il client usa la regione Firebase corretta");
assert.doesNotMatch(client, /const admin = auth\.currentUser/, "La gestione admin non dipende da auth non inizializzato");
assert.match(client, /const admin = firebase\.auth\(\)\.currentUser/, "La gestione admin legge la sessione Firebase corrente");
assert.match(client, /batch\.update\(ref, patch\)/, "Il fallback conserva lo sblocco Firestore compatibile");
assert.match(client, /approval_\$\{uid\}_\$\{requestId\}/, "Il fallback usa un audit idempotente");
assert.match(client, /Google Play/, "Il messaggio spiega l’installazione Android");
assert.match(client, /Aggiungi alla schermata Home/, "Il messaggio spiega l’installazione iPhone");
assert.match(client, /password che hai scelto durante la registrazione/, "Il messaggio indica le credenziali scelte dall’utente");

assert.match(backend, /exports\.approveUserAccess/, "Il backend esporta approveUserAccess");
assert.match(backend, /getAdminEmails/, "Il backend verifica l’amministratore");
assert.match(backend, /batch\.update\(userRef, approvalPatch\)/, "Il backend autorizza il profilo senza sostituirlo");
assert.match(backend, /RESEND_ENDPOINT/, "Il backend invia l’email tramite il servizio configurato");
assert.match(backend, /optionalEmailConfiguration/, "Il backend distribuisce anche senza segreti email mancanti");
assert.doesNotMatch(backend, /defineSecret|runWith\(\{ secrets:/, "Il deploy non richiede segreti interattivi assenti");
assert.match(backend, /Idempotency-Key/, "L’email non viene duplicata durante i tentativi");
assert.match(backend, /emailError = "Invio email non riuscito/, "Un errore email non annulla lo sblocco");
assert.doesNotMatch(backend, /\.delete\(|batch\.delete\(/, "Il flusso non cancella utenti o collegamenti");
assert.match(entrypoint, /user-access-approval/, "Il callable è incluso nell’entrypoint Firebase");
assert.match(workflow, /functions:approveUserAccess/, "Il workflow distribuisce il callable");
assert.match(workflow, /"approveUserAccess"/, "Il workflow verifica il trasporto pubblico del callable");
assert.match(index, /approval-access\.js\?v=20260828-authfix/, "La pagina carica la correzione Auth più recente");
assert.match(serviceWorker, /approval-access\.js\?v=20260828-authfix/, "La PWA precarica la correzione Auth più recente");

console.log("OK: sblocco utente, email automatica e messaggio WhatsApp verificati.");

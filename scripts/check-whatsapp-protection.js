#!/usr/bin/env node
'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const protectedFile = 'android/app/src/main/java/it/vargacantieri/hera/whatsapp/HeraWhatsAppPlugin.java';
const expectedGitBlobSha = 'c7145043e56dc98984827c149c6d9fce46bf6d8b';

function fail(message) {
  console.error(`\n❌ PROTEZIONE WHATSAPP BLOCCATA\n${message}\n`);
  process.exit(1);
}

function gitBlobSha(buffer) {
  const header = Buffer.from(`blob ${buffer.length}\0`, 'utf8');
  return crypto.createHash('sha1').update(Buffer.concat([header, buffer])).digest('hex');
}

const absolutePath = path.join(root, protectedFile);
if (!fs.existsSync(absolutePath)) fail(`File protetto mancante: ${protectedFile}`);

const buffer = fs.readFileSync(absolutePath);
const source = buffer.toString('utf8');
const actualSha = gitBlobSha(buffer);

if (actualSha !== expectedGitBlobSha) {
  fail([
    `Il file protetto è stato modificato: ${protectedFile}`,
    `Hash atteso: ${expectedGitBlobSha}`,
    `Hash trovato: ${actualSha}`,
    'Ripristinare la versione approvata. Qualsiasi modifica richiede autorizzazione esplicita e aggiornamento consapevole di questo controllo.'
  ].join('\n'));
}

const requiredMarkers = [
  '@CapacitorPlugin(name = "HeraWhatsApp")',
  'private static final String WHATSAPP = "com.whatsapp";',
  'private static final String WHATSAPP_BUSINESS = "com.whatsapp.w4b";',
  'String packageName = resolveInstalledPackage();',
  'if (packageName == null)',
  'intent.setPackage(packageName);',
  'intent.putExtra(Intent.EXTRA_TEXT, payload.text);',
  'intent.putExtra("jid", payload.phone + "@s.whatsapp.net");',
  'openWebFallback(call, rawUrl, payload, "Impossibile aprire WhatsApp installato.");',
  'https://api.whatsapp.com/send',
  'https://wa.me/'
];

for (const marker of requiredMarkers) {
  if (!source.includes(marker)) fail(`Comportamento WhatsApp protetto non trovato: ${marker}`);
}

const appFirst = source.indexOf('resolveInstalledPackage()');
const webFallback = source.indexOf('openWebFallback(call, rawUrl, payload');
if (appFirst < 0 || webFallback < 0 || appFirst > webFallback) {
  fail('La priorità app WhatsApp installata → fallback web non è più garantita.');
}

if (/openWebFallback\([^)]*\);\s*openWebFallback\(/s.test(source)) {
  fail('Rilevato possibile fallback WhatsApp duplicato.');
}

console.log('✅ WhatsApp/WHAZZUP protetto: file immutato e regole critiche verificate.');
console.log(`✅ Impronta approvata: ${expectedGitBlobSha}`);

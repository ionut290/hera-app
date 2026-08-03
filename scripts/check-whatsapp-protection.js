#!/usr/bin/env node
'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const protectedFile = 'android/app/src/main/java/it/vargacantieri/hera/whatsapp/HeraWhatsAppPlugin.java';
const expectedGitBlobSha = 'c1bf032d116e2a7adf32413e103872272ee848a8';

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
    'Ripristinare la versione approvata. La logica WhatsApp attuale non deve essere modificata.'
  ].join('\n'));
}

const requiredMarkers = [
  '@CapacitorPlugin(name = "HeraWhatsApp")',
  'private static final String WHATSAPP = "com.whatsapp";',
  'private static final String WHATSAPP_BUSINESS = "com.whatsapp.w4b";',
  'String packageName = resolveInstalledPackage();',
  'call.reject("WhatsApp non è installato sul dispositivo.");',
  'intent.setPackage(packageName);',
  'intent.putExtra(Intent.EXTRA_TEXT, payload.text);',
  'intent.putExtra("jid", payload.phone + "@s.whatsapp.net");',
  'call.reject("Impossibile aprire WhatsApp installato.", error);',
  'result.put("fallback", false);'
];

for (const marker of requiredMarkers) {
  if (!source.includes(marker)) fail(`Comportamento WhatsApp protetto non trovato: ${marker}`);
}

const forbiddenMarkers = [
  'openWebFallback',
  'buildWebUrl',
  'Intent.ACTION_VIEW',
  'browserIntent',
  'result.put("fallback", true)',
  'Impossibile aprire WhatsApp installato o WhatsApp Web'
];

for (const marker of forbiddenMarkers) {
  if (source.includes(marker)) fail(`WhatsApp Web o fallback browser reintrodotto: ${marker}`);
}

const packageResolution = source.indexOf('resolveInstalledPackage()');
const nativeSend = source.indexOf('new Intent(Intent.ACTION_SEND)');
if (packageResolution < 0 || nativeSend < 0 || packageResolution > nativeSend) {
  fail('L’apertura dell’app WhatsApp installata non è più garantita.');
}

console.log('✅ WhatsApp/WHAZZUP protetto: solo app installata, nessun fallback web.');
console.log(`✅ Impronta approvata: ${expectedGitBlobSha}`);

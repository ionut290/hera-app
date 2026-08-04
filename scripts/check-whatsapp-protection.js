#!/usr/bin/env node
'use strict';

const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const protectedFiles = [
  {
    path: 'android/app/src/main/java/it/vargacantieri/hera/whatsapp/HeraWhatsAppPlugin.java',
    sha: 'c1bf032d116e2a7adf32413e103872272ee848a8'
  },
  {
    path: 'whazzup-preload-cache.js',
    sha: 'd27ca8a36ae5a32ae9e90c9c572cb7d37c491fad'
  }
];

function fail(message) {
  console.error(`\n❌ PROTEZIONE WHATSAPP BLOCCATA\n${message}\n`);
  process.exit(1);
}

function gitBlobSha(buffer) {
  const header = Buffer.from(`blob ${buffer.length}\0`, 'utf8');
  return crypto.createHash('sha1').update(Buffer.concat([header, buffer])).digest('hex');
}

function readFile(relativePath) {
  const absolutePath = path.join(root, relativePath);
  if (!fs.existsSync(absolutePath)) fail(`File mancante: ${relativePath}`);
  return fs.readFileSync(absolutePath, 'utf8');
}

function readProtectedFile(definition) {
  const absolutePath = path.join(root, definition.path);
  if (!fs.existsSync(absolutePath)) fail(`File protetto mancante: ${definition.path}`);
  const buffer = fs.readFileSync(absolutePath);
  const actualSha = gitBlobSha(buffer);
  if (actualSha !== definition.sha) {
    fail([
      `Il file protetto è stato modificato: ${definition.path}`,
      `Hash atteso: ${definition.sha}`,
      `Hash trovato: ${actualSha}`,
      'Ripristinare la versione approvata. La logica WhatsApp attuale non deve essere modificata.'
    ].join('\n'));
  }
  return buffer.toString('utf8');
}

const nativeSource = readProtectedFile(protectedFiles[0]);
const webSource = readProtectedFile(protectedFiles[1]);
const loaderSource = readFile('varga-branding.js');
const serviceWorkerSource = readFile('sw.js');
const appSource = readFile('app.js');

const nativeRequiredMarkers = [
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

for (const marker of nativeRequiredMarkers) {
  if (!nativeSource.includes(marker)) fail(`Comportamento WhatsApp nativo protetto non trovato: ${marker}`);
}

const nativeForbiddenMarkers = [
  'openWebFallback',
  'buildWebUrl',
  'Intent.ACTION_VIEW',
  'browserIntent',
  'result.put("fallback", true)',
  'Impossibile aprire WhatsApp installato o WhatsApp Web'
];

for (const marker of nativeForbiddenMarkers) {
  if (nativeSource.includes(marker)) fail(`Fallback browser Android reintrodotto: ${marker}`);
}

const webRequiredMarkers = [
  'installWhatsAppInstalledAppOnlyOpen',
  'return `whatsapp://send?${params.toString()}`;',
  'window.location.assign(directUrl);',
  'WhatsApp non è installato o non può essere aperto su questo dispositivo.'
];

for (const marker of webRequiredMarkers) {
  if (!webSource.includes(marker)) fail(`Comportamento WhatsApp PWA protetto non trovato: ${marker}`);
}

const webForbiddenMarkers = [
  'openDirectWithFallback',
  'Vuoi aprire WhatsApp Web?',
  'window.location.assign(webUrl)',
  'window.open(webUrl',
  'location.href = webUrl'
];

for (const marker of webForbiddenMarkers) {
  if (webSource.includes(marker)) fail(`Fallback WhatsApp Web reintrodotto nella PWA: ${marker}`);
}

const appForbiddenFallbackMarkers = [
  'window.location.href = webUrl',
  'targetWindow.location.replace(webUrl)',
  'if (!opened && !disableWebFallback)',
  'Errore fallback WhatsApp wa.me:'
];

for (const marker of appForbiddenFallbackMarkers) {
  if (appSource.includes(marker)) fail(`Fallback WhatsApp Web reintrodotto nel flusso principale: ${marker}`);
}

const loaderMarkers = [
  'function loadWhatsAppInstalledOnlyGuard()',
  './whazzup-preload-cache.js?v=20260803-installed-only2',
  "'whazzupInstalledOnlyGuard'",
  'loadWhatsAppInstalledOnlyGuard();'
];
for (const marker of loaderMarkers) {
  if (!loaderSource.includes(marker)) fail(`Il blocco WhatsApp PWA non viene caricato all’avvio: ${marker}`);
}

const cacheVersionMatch = serviceWorkerSource.match(/varga-cantieri-shell-v(\d+)/);
const cacheVersion = Number(cacheVersionMatch?.[1] || 0);
if (cacheVersion < 94) {
  fail('La cache PWA non è aggiornata alla versione che rimuove il vecchio fallback WhatsApp Web.');
}
if (!serviceWorkerSource.includes('./whazzup-preload-cache.js?v=20260803-installed-only2')) {
  fail('Il modulo WhatsApp installato-only non è presente nella cache PWA.');
}

console.log('✅ WhatsApp/WHAZZUP protetto su Android e PWA: solo app installata, nessun fallback web.');
console.log(`✅ Il blocco PWA viene caricato all’avvio ed è incluso nella cache corrente v${cacheVersion}.`);
protectedFiles.forEach((file) => console.log(`✅ ${file.path}: ${file.sha}`));

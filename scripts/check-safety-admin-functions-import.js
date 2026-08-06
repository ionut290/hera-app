const fs = require('fs');
const path = require('path');

const root = process.cwd();
const read = (file) => fs.readFileSync(path.join(root, file), 'utf8');
const existsNonEmpty = (file) => {
  const full = path.join(root, file);
  return fs.existsSync(full) && fs.statSync(full).size > 0;
};
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
  console.log(`✓ ${message}`);
};

const requiredFiles = [
  'firestore.rules',
  'firebase.json',
  'functions/main.js',
  'functions/package.json',
  'admin-console.js',
  'notification-session-enhancements.js',
  'registry-google-sheet-sync.js',
  'google-sheet-two-way-sync.js',
  'rubrica-matrice-personale-import.js',
  'sw.js'
];

for (const file of requiredFiles) {
  assert(existsNonEmpty(file), `${file} esiste e non è vuoto`);
}

const rules = read('firestore.rules');
const firebaseJson = JSON.parse(read('firebase.json'));
const functionsPackage = JSON.parse(read('functions/package.json'));
const functionsMain = read('functions/main.js');
const adminConsole = read('admin-console.js');
const notifications = read('notification-session-enhancements.js');
const registrySync = read('registry-google-sheet-sync.js');
const sheetSync = read('google-sheet-two-way-sync.js');
const excelImport = read('rubrica-matrice-personale-import.js');
const sw = read('sw.js');
const app = read('app.js');
const index = read('index.html');

// Regole Firestore e separazione dei privilegi.
assert(/rules_version\s*=\s*['"]2['"]/.test(rules), 'Le regole Firestore usano rules_version 2');
assert(/function\s+signedIn\s*\(/.test(rules), 'Le regole definiscono il controllo signedIn');
assert(/function\s+isAdmin\s*\(/.test(rules), 'Le regole definiscono il controllo isAdmin');
assert(/match\s+\/platformUsers\//.test(rules), 'Le regole proteggono i profili utente');
assert(/match\s+\/operatorPositions\//.test(rules) && /allow\s+read:\s*if\s+isAdmin\(\)/.test(rules), 'Le posizioni operatori restano leggibili soltanto dagli admin');
assert(/match\s+\/activityLogs\//.test(rules), 'Le regole proteggono i log attività');
assert(/match\s+\/privateDocuments\//.test(rules) && /request\.auth\.uid\s*==\s*userId/.test(rules), 'I documenti privati restano limitati al proprietario');

// Firebase e Cloud Functions.
assert(firebaseJson && typeof firebaseJson === 'object', 'firebase.json è valido');
assert(firebaseJson.functions || firebaseJson.firestore, 'firebase.json collega Functions o Firestore');
assert(functionsPackage.dependencies && Object.keys(functionsPackage.dependencies).some((name) => /firebase-functions/.test(name)), 'Functions dichiara firebase-functions');
assert(/exports\.|module\.exports|onCall|onRequest|onDocument|functions\./.test(functionsMain), 'functions/main.js espone funzioni Cloud');
assert(!/BEGIN PRIVATE KEY|PRIVATE KEY-----/.test(functionsMain), 'Nessuna chiave privata è incorporata nelle Cloud Functions');

// Permessi amministrativi lato interfaccia.
assert(/admin/i.test(adminConsole), 'Il modulo console amministrativa contiene i controlli admin');
assert(/isAdmin|adminOnly|requireAdmin|currentUser.*admin/i.test(adminConsole + app), 'La logica applicativa verifica il ruolo amministratore');

// Notifiche e comportamento offline/PWA.
assert(/notification|notifica/i.test(notifications), 'Il modulo notifiche è presente');
assert(/self\.addEventListener\(['"]install['"]/.test(sw), 'Il Service Worker gestisce install');
assert(/self\.addEventListener\(['"]activate['"]/.test(sw), 'Il Service Worker gestisce activate');
assert(/self\.addEventListener\(['"]fetch['"]/.test(sw), 'Il Service Worker gestisce fetch');
assert(/navigator\.onLine|['"]online['"]|['"]offline['"]/.test(app + index + sw), 'L’app contiene una gestione online/offline');

// Import Excel e sincronizzazione Google Sheets.
assert(/XLSX|FileReader|arrayBuffer/i.test(excelImport), 'L’import del registro gestisce file Excel');
assert(/sheet|foglio|spreadsheet/i.test(registrySync), 'La sincronizzazione registro fa riferimento a Google Sheets');
assert(/sheet|foglio|spreadsheet/i.test(sheetSync), 'La sincronizzazione bidirezionale fa riferimento ai fogli');
assert(/fetch\s*\(|XMLHttpRequest|googleapis|script\.google/i.test(registrySync + sheetSync), 'La sincronizzazione usa un canale di comunicazione esplicito');

// Nessun segreto server-side nei principali file client controllati.
const clientBundle = [app, index, adminConsole, notifications, registrySync, sheetSync, excelImport].join('\n');
assert(!/BEGIN PRIVATE KEY|PRIVATE KEY-----|service_account\s*[:=]/i.test(clientBundle), 'Nessuna chiave privata server-side è esposta nei file client controllati');

console.log('\nControlli admin, Functions, import, notifiche e offline completati senza connessioni esterne.');

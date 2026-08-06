const fs = require('fs');
const path = require('path');

const read = (file) => fs.readFileSync(path.join(process.cwd(), file), 'utf8');
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
  console.log(`✓ ${message}`);
};

const index = read('index.html');
const sw = read('sw.js');
const manifest = JSON.parse(read('manifest.webmanifest'));
const firebaseConfig = read('firebase-config.js');
const app = read('app.js');

// Gestione ore: protegge la presenza del comando e dei contenitori principali.
assert(/id=["']open-hours-btn["']/.test(index), 'Il pulsante Gestione ore esiste');
assert(/Gestione ore/i.test(index), 'L’etichetta Gestione ore è presente');

// Mappa: verifica che Leaflet e almeno un contenitore/azione mappa restino disponibili.
assert(/leaflet@1\.9\.4/i.test(index), 'Leaflet è caricato dalla pagina');
assert(/id=["'][^"']*map[^"']*["']/i.test(index), 'È presente almeno un elemento mappa');
assert(/google\.com\/maps|maps\.google|openstreetmap|tile\.openstreetmap/i.test(index + app), 'La navigazione o la cartografia è referenziata');

// PWA e Service Worker.
assert(/rel=["']manifest["'][^>]*manifest\.webmanifest/i.test(index), 'Il manifest PWA è collegato');
assert(/serviceWorker\.register\(["']\.\/sw\.js["']\)/.test(index), 'Il Service Worker viene registrato');
assert(typeof manifest.name === 'string' && manifest.name.trim(), 'Il manifest contiene il nome dell’app');
assert(typeof manifest.start_url === 'string' && manifest.start_url.trim(), 'Il manifest contiene start_url');
assert(Array.isArray(manifest.icons) && manifest.icons.length >= 2, 'Il manifest contiene le icone principali');
assert(/self\.addEventListener\(["']install["']/.test(sw), 'Il Service Worker gestisce install');
assert(/self\.addEventListener\(["']activate["']/.test(sw), 'Il Service Worker gestisce activate');
assert(/self\.addEventListener\(["']fetch["']/.test(sw), 'Il Service Worker gestisce fetch');

const appVersionIndex = index.match(/app\.js\?v=([^"']+)/)?.[1];
const appVersionSw = sw.match(/app\.js\?v=([^"']+)/)?.[1];
assert(Boolean(appVersionIndex), 'index.html usa una versione esplicita di app.js');
assert(appVersionIndex === appVersionSw, 'index.html e sw.js usano la stessa versione di app.js');

// Firebase/Firestore: controlli statici e non invasivi, senza connessione o scritture.
assert(/initializeApp|firebaseConfig|firebase\.initializeApp/i.test(firebaseConfig), 'Firebase è inizializzato');
assert(/firestore/i.test(firebaseConfig + app), 'Firestore è referenziato dall’app');
assert(!/BEGIN PRIVATE KEY|PRIVATE KEY-----|service_account/i.test(firebaseConfig), 'Non sono presenti chiavi private di servizio nel client');

const optimizerFiles = [
  'firestore-safe-optimizer.js',
  'firestore-inflight-read-coalescer.js',
  'firestore-nested-listener-optimizer.js'
];
for (const file of optimizerFiles) {
  assert(fs.existsSync(path.join(process.cwd(), file)), `Esiste la protezione ${file}`);
  assert(index.includes(file), `${file} è caricato da index.html`);
}

console.log('\nControlli estesi completati senza collegarsi a Firestore o modificare dati.');

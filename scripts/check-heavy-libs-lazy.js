const fs = require('fs');

function fail(message) {
  console.error(`❌ ${message}`);
  process.exit(1);
}

const index = fs.readFileSync('index.html', 'utf8');
const sw = fs.readFileSync('sw.js', 'utf8');
const loader = fs.readFileSync('heavy-libs-lazy-loader.js', 'utf8');
const hours = fs.readFileSync('hours-export-range.js', 'utf8');
const identity = fs.readFileSync('identity-card-feature.js', 'utf8');

const removed = [
  'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
  'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js',
  'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js'
];
for (const src of removed) {
  if (index.includes(src)) fail(`libreria pesante ancora caricata all'avvio: ${src}`);
  if (!loader.includes(src)) fail(`loader lazy non conosce la libreria: ${src}`);
}

const lazyTag = 'heavy-libs-lazy-loader.js?v=20260815a';
if (!index.includes(lazyTag)) fail('loader lazy non incluso in index.html');
if (!sw.includes(`./${lazyTag}`)) fail('loader lazy non incluso nell app shell offline');

const lazyPos = index.indexOf(lazyTag);
const firebaseConfigPos = index.indexOf('firebase-config.js');
const appPos = index.indexOf('app.js?v=');
if (!(lazyPos >= 0 && firebaseConfigPos > lazyPos && appPos > lazyPos)) {
  fail('ordine loader lazy non sicuro rispetto a firebase-config/app.js');
}

if (!index.includes('https://unpkg.com/leaflet@1.9.4/dist/leaflet.js')) {
  fail('Leaflet è stato rimosso dalla fase 1: rischio mappa non consentito');
}

if (!hours.includes('HeraHeavyLibs.ensure("exceljs")')) {
  fail('export ore non ha fallback diretto per ExcelJS lazy');
}
if (!identity.includes('HeraHeavyLibs.ensure("html2canvas")')) {
  fail('carta da visita non ha fallback diretto per html2canvas lazy');
}

const firestorePatterns = [
  /firebase\s*\.\s*firestore\b/,
  /\bdb\s*\.\s*collection\s*\(/,
  /\.onSnapshot\s*\(/,
  /firebase\s*\.\s*database\b/,
  /\bgetFirestore\s*\(/
];
for (const pattern of firestorePatterns) {
  if (pattern.test(loader)) fail(`loader lazy non deve eseguire Firestore: ${pattern}`);
}
for (const protectedTerm of ['fattoVisualEvidence', 'WHAZZUP', 'WhatsApp', 'completeImpianto', 'markImpiantoDone']) {
  if (loader.includes(protectedTerm)) fail(`loader lazy contiene riferimento a flusso protetto: ${protectedTerm}`);
}

console.log('✅ Lazy heavy libs: avvio alleggerito, fallback presenti, Leaflet e flussi protetti invariati.');

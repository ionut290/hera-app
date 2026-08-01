'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const cp = require('child_process');

const root = path.join(__dirname, '..');
const files = {
  export: path.join(root, 'preventivi-models-export.js'),
  documents: path.join(root, 'preventivi-models-documents.js'),
  consuntivi: path.join(root, 'preventivi-consuntivi.js'),
  exact: path.join(root, 'preventivi-exact-xlsx.js'),
  followup: path.join(root, 'preventivi-registry-model-followup.js'),
  header: path.join(root, 'header-menu-runtime.js'),
  sw: path.join(root, 'sw.js')
};

Object.values(files).forEach((file) => cp.execFileSync(process.execPath, ['--check', file], { stdio: 'pipe' }));

const source = Object.fromEntries(Object.entries(files).map(([name, file]) => [name, fs.readFileSync(file, 'utf8')]));

[
  'prezzo_capitolato',
  'ribasso_si_no',
  'prezzo_netto_ribassato',
  'importo_prestazione',
  'note_attivita'
].forEach((token) => assert(source.export.includes(token), `Dato economico mancante: ${token}`));

[
  'data-price-list-id',
  'data-price-item-id',
  'data-contract-price',
  'data-discount',
  'collectConsLines'
].forEach((token) => assert(source.documents.includes(token), `Collegamento prezziario mancante: ${token}`));

assert(source.documents.includes('PV.priceItemNetPrice?.(item)'), 'Il consuntivo non usa il prezzo netto del prezziario.');
assert(source.consuntivi.includes('data-contract-price'), 'La riapertura del consuntivo perde il prezzo di capitolato.');
assert(source.consuntivi.includes('priceItemId:String(row.dataset.priceItemId'), 'Il salvataggio del consuntivo perde il collegamento alla voce prezziario.');
assert(source.exact.includes('const uniqueCols=[...new Set(Object.values(cols).filter(Boolean))]'), 'Le colonne non riconosciute non sono escluse.');
assert(source.exact.includes('const tableEnd=total?pos(total.ref).row-1'), 'Il limite della tabella lavorazioni non è protetto.');
assert(!source.exact.includes('fallback)=>pos((find(heads,labels)||fallback)'), 'È ancora presente il riuso pericoloso di una colonna non riconosciuta.');
assert(source.followup.includes('preventivi-exact-xlsx.js?v=20260801b'), 'Il compilatore XLSX aggiornato non viene caricato.');
assert(source.header.includes('preventivi-models-documents.js?v=20260801c'), 'Il modulo documenti aggiornato non viene caricato.');
assert(source.header.includes('preventivi-models-export.js?v=20260801c'), 'Il modulo esportazione aggiornato non viene caricato.');
assert(source.sw.includes('varga-cantieri-shell-v85'), 'Cache PWA non aggiornata.');
assert(source.sw.includes('preventivi-exact-xlsx.js?v=20260801b'), 'Compilatore XLSX aggiornato assente dalla cache.');

console.log('check-preventivi-template-fidelity: OK');

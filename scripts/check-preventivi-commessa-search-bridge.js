'use strict';
const assert = require('node:assert');
const fs = require('node:fs');
const cp = require('node:child_process');

const bridge = fs.readFileSync('preventivi-commessa-search-bridge.js', 'utf8');
const header = fs.readFileSync('header-menu-runtime.js', 'utf8');
const sw = fs.readFileSync('sw.js', 'utf8');

cp.execFileSync(process.execPath, ['--check', 'preventivi-commessa-search-bridge.js']);

[
  "collection('globalArchive').doc('commesse').collection('items')",
  "for (const collectionName of ['impiantiFisici', 'impianti'])",
  'Cerca impianto, comune, indirizzo, ID SAP o codice prezzo',
  "group.label = 'GLOBAL'",
  'selectionByEditor: new Map()',
  'function rememberSelection(form, select)',
  'function rememberedSelection(form)',
  'function restoreSelection(form, select)',
  "restoreSelection(form, form?.querySelector(COMMESSA_SELECTOR))",
  "setBackingValue(form, 'commessaSource', saved.scope)",
  'data-commessa-plant',
  "version: '20260801b'"
].forEach((token) => assert(bridge.includes(token), `Manca nel bridge: ${token}`));

assert(header.includes("['commessa-search-bridge', './preventivi-commessa-search-bridge.js?v=20260801a']"), 'Bridge non caricato dal runtime.');
assert(header.indexOf("['commessa-search-bridge'") > header.indexOf("['draft-preserver'"), 'Il bridge deve essere caricato dopo la protezione della bozza.');
assert(sw.includes('varga-cantieri-shell-v84'), 'Cache PWA non aggiornata a v84.');
assert(sw.includes('./preventivi-commessa-search-bridge.js?v=20260801a'), 'Bridge non presente nella cache PWA.');

console.log('check-preventivi-commessa-search-bridge: OK');

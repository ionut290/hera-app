'use strict';
const assert=require('node:assert');
const fs=require('node:fs');
const cp=require('node:child_process');
const bridge=fs.readFileSync('preventivi-commessa-search-bridge.js','utf8');
const header=fs.readFileSync('header-menu-runtime.js','utf8');
const sw=fs.readFileSync('sw.js','utf8');
cp.execFileSync(process.execPath,['--check','preventivi-commessa-search-bridge.js']);
[
  "collection('globalArchive').doc('commesse').collection('items')",
  "for(const n of ['impiantiFisici','impianti'])",
  'Cerca impianto, comune, indirizzo, ID SAP o codice prezzo',
  "group.label='GLOBAL'",
  "value=\"global::${E(x.id)}\"",
  '[data-matrix-plant-search], [data-pvd-plant-search]',
  'data-commessa-plant',
  "backing(form,'commessaSource').value=c.scope"
].forEach(token=>assert(bridge.includes(token),`Manca nel bridge: ${token}`));
assert(header.includes("['commessa-search-bridge', './preventivi-commessa-search-bridge.js?v=20260801a']"),'Bridge non caricato dal runtime.');
assert(header.indexOf("['commessa-search-bridge'")>header.indexOf("['matrix-runtime-fix'"),'Il bridge deve essere caricato dopo la hotfix matrice.');
assert(sw.includes('varga-cantieri-shell-v83'),'Cache PWA non aggiornata a v83.');
assert(sw.includes('./preventivi-commessa-search-bridge.js?v=20260801a'),'Bridge non presente nella cache PWA.');
console.log('check-preventivi-commessa-search-bridge: OK');

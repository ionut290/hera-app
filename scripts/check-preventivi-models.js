'use strict';
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const names = ['preventivi-models-core.js','preventivi-models-ui.js','preventivi-models-documents.js','preventivi-models-export.js'];
const source = names.map(name => fs.readFileSync(path.join(root, name), 'utf8')).join('\n');
const css = fs.readFileSync(path.join(root, 'preventivi-models.css'), 'utf8');
const loader = fs.readFileSync(path.join(root, 'header-menu-runtime.js'), 'utf8');
const sw = fs.readFileSync(path.join(root, 'sw.js'), 'utf8');
[
  'Modelli preventivi e consuntivi', 'Prezziario di riferimento', 'Campi riconosciuti nel modello',
  'exportDocx', 'exportSheet', 'exportFillablePdf', 'data-pvm-dynamic-fields',
  'commesse/preventivi_app/modelli', 'commesse/preventivi_app/consuntivi'
].forEach(needle => assert(source.includes(needle), `Manca: ${needle}`));
[
  "['models-core', './preventivi-models-core.js?v=20260801b']",
  "['models-ui', './preventivi-models-ui.js?v=20260801b']",
  "['models-documents', './preventivi-models-documents.js?v=20260801b']",
  "['models-export', './preventivi-models-export.js?v=20260801b']"
].forEach(needle => assert(loader.includes(needle), `Loader mancante: ${needle}`));
assert(loader.includes('preventivi-models.css?v=20260801a'), 'CSS modelli non caricato.');
assert(sw.includes('varga-cantieri-shell-v73'), 'Cache PWA non aggiornata.');
names.forEach(name => assert(sw.includes(name), `File non incluso nella cache: ${name}`));
assert(sw.includes('preventivi-models.css?v=20260801a'), 'CSS non incluso nella cache.');
assert(css.includes('.pvm-dynamic-fields'), 'Stili campi dinamici mancanti.');
console.log('check-preventivi-models: OK');

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
const draft = fs.readFileSync(path.join(root, 'preventivi-draft-preserver.js'), 'utf8');
[
  'Modelli preventivi e consuntivi', 'Prezziario di riferimento', 'Campi riconosciuti nel modello',
  'exportDocx', 'exportSheet', 'exportFillablePdf', 'data-pvm-dynamic-fields',
  'commesse/preventivi_app/modelli', 'commesse/preventivi_app/consuntivi'
].forEach(needle => assert(source.includes(needle), `Manca: ${needle}`));
[
  "['models-core', './preventivi-models-core.js?v=20260801b']",
  "['models-ui', './preventivi-models-ui.js?v=20260801b']",
  "['models-documents', './preventivi-models-documents.js?v=20260801d']",
  "['models-export', './preventivi-models-export.js?v=20260801c']"
].forEach(needle => assert(loader.includes(needle), `Loader mancante: ${needle}`));
assert(loader.includes('preventivi-models.css?v=20260801a'), 'CSS modelli non caricato.');
assert(sw.includes('varga-cantieri-shell-v86'), 'Cache PWA non aggiornata.');
assert(source.includes('M.liveModelFields'), 'La ricostruzione dei campi non conserva i valori digitati.');
assert(loader.includes("['draft-preserver', './preventivi-draft-preserver.js?v=20260801b']"), 'Protezione bozza aggiornata non caricata.');
assert(draft.includes("saved.key === 'modelId'"), 'Il modello scelto non viene ripristinato prima dei campi.');
assert(draft.includes('HeraPreventiviModels?.renderDynamic'), 'I campi del modello scelto non vengono ricostruiti durante il ripristino.');
assert(draft.includes("saved.key.startsWith('control-')"), 'Il ripristino può ancora scrivere dati in campi diversi.');
names.forEach(name => assert(sw.includes(name), `File non incluso nella cache: ${name}`));
assert(sw.includes('preventivi-models.css?v=20260801a'), 'CSS non incluso nella cache.');
assert(css.includes('.pvm-dynamic-fields'), 'Stili campi dinamici mancanti.');
console.log('check-preventivi-models: OK');

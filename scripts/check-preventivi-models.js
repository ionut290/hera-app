'use strict';

const assert = require('assert');
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const source = fs.readFileSync(path.join(root, 'preventivi-models.js'), 'utf8');
const css = fs.readFileSync(path.join(root, 'preventivi-models.css'), 'utf8');
const loader = fs.readFileSync(path.join(root, 'header-menu-runtime.js'), 'utf8');
const sw = fs.readFileSync(path.join(root, 'sw.js'), 'utf8');

[
  'Modelli preventivi e consuntivi',
  'Prezziario di riferimento',
  'Campi riconosciuti nel modello',
  'exportDocx',
  'exportSpreadsheet',
  'exportFillablePdf',
  'data-pvm-dynamic-fields',
  'commesse/preventivi_app/modelli',
  'commesse/preventivi_app/consuntivi'
].forEach((needle) => assert(source.includes(needle), `Manca: ${needle}`));

assert(loader.includes("['models', './preventivi-models.js?v=20260801a']"), 'Il modulo modelli non è caricato in sequenza.');
assert(loader.includes('preventivi-models.css?v=20260801a'), 'Il CSS modelli non è caricato.');
assert(sw.includes('varga-cantieri-shell-v73'), 'La cache PWA non è stata aggiornata.');
assert(sw.includes('preventivi-models.js?v=20260801a'), 'Il modulo non è nella cache PWA.');
assert(sw.includes('preventivi-models.css?v=20260801a'), 'Il CSS non è nella cache PWA.');
assert(css.includes('.pvm-dynamic-fields'), 'Manca lo stile dei campi dinamici.');

console.log('check-preventivi-models: OK');

const fs = require('node:fs');

const read = (path) => fs.readFileSync(path, 'utf8');
const runtime = read('preventivi-matrix-runtime-fix.js');
const header = read('preventivi-lazy-loader.js');
const sw = read('sw.js');

const checks = [
  [runtime.includes("form.querySelectorAll('[data-pvd-section]').forEach(section=>section.remove())"), 'rimozione sezione modello duplicata'],
  [runtime.includes('[data-matrix-runtime-placeholder]'), 'protezione campi dinamici duplicati'],
  [runtime.includes("typeof impiantiById!=='undefined'"), 'lettura archivio impianti applicazione'],
  [runtime.includes("readStorage(window.localStorage,'localStorage')"), 'fallback archivio locale'],
  [runtime.includes('[data-matrix-runtime-plant]'), 'selezione impianto riparata'],
  [runtime.includes('Il collegamento automatico con la commessa non era presente'), 'fallback commessa senza relazione esplicita'],
  [header.includes("['matrix-runtime-fix', './preventivi-matrix-runtime-fix.js?v=20260801a']"), 'caricamento runtime hotfix'],
  [header.indexOf("['matrix-runtime-fix'") > header.indexOf("['tabs-models-plant-guard'"), 'ordine caricamento dopo guardia modelli'],
  [sw.includes('varga-cantieri-shell-v114'), 'cache PWA corrente'],
  [header.includes('./preventivi-matrix-runtime-fix.js?v=20260801a'), 'hotfix nel caricamento su richiesta']
];

const failed = checks.filter(([ok]) => !ok).map(([, label]) => label);
if (failed.length) {
  console.error(`check-preventivi-matrix-runtime-fix: FAIL\n- ${failed.join('\n- ')}`);
  process.exit(1);
}
console.log('check-preventivi-matrix-runtime-fix: OK');

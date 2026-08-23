const fs = require('fs');
const path = require('path');

const root = process.argv[2] || path.resolve(__dirname, '..');
const read = (name) => fs.readFileSync(path.join(root, name), 'utf8');
const core = read('recommended-plants.js');
const css = read('recommended-plants.css');
const adaptive = read('adaptive-work-learning.js');
const equipment = read('equipment-recommendations.js');
const guard = read('fatto-scroll-guard.js');
const sw = read('sw.js');

const checks = [
  [!adaptive.includes('new MutationObserver'), 'adaptive-work-learning non deve creare observer ricorsivi'],
  [!equipment.includes('new MutationObserver'), 'equipment-recommendations non deve creare observer ricorsivi'],
  [!guard.includes('equipment-recommendations.js'), 'il lazy loader non deve ricaricare il vecchio motore attrezzature'],
  [!guard.includes('adaptive-work-learning.js'), 'il lazy loader non deve ricaricare il vecchio motore adattivo'],
  [core.includes('out.length<CFG.max') || core.includes('plan.length < CONFIG.maxVisible'), 'il percorso deve fermarsi agli impianti visibili della giornata'],
  [core.includes('await sleep0()') || core.includes('await yieldToUi()'), 'il calcolo deve cedere il controllo alla UI'],
  [!core.includes('remaining.sort('), 'il calcolo non deve riordinare ripetutamente tutti gli impianti'],
  [core.includes('function close()') || core.includes('closePanel()'), 'deve esistere la chiusura immediata del pannello'],
  [core.includes('state.seq++') || core.includes('state.renderSequence += 1'), 'la chiusura deve annullare il calcolo ancora in corso'],
  [!core.includes('childList:true') && !core.includes('subtree:true'), 'il motore non deve osservare le mutazioni interne della pagina Impianti'],
  [!css.match(/z-index\s*:\s*40/i), 'il pannello non deve usare z-index 40'],
  [!css.match(/isolation\s*:\s*isolate/i), 'il pannello non deve creare uno stacking context isolato'],
  [css.includes('.recommended-plants-panel.hidden'), 'il pannello nascosto deve essere realmente rimosso dai tocchi'],
  [sw.includes('varga-cantieri-shell-v138'), 'il service worker deve usare una cache nuova'],
  [sw.includes('CACHE_RESET_VERSION = "20260814-loading-humor1"'), 'la versione reset cache deve restare allineata al bootstrap PWA'],
  [sw.includes('"/recommended-plants.js"'), 'recommended-plants.js deve essere network-first'],
  [sw.includes('"/recommended-plants.css"'), 'recommended-plants.css deve essere network-first'],
  [sw.includes('"/fatto-scroll-guard.js"'), 'fatto-scroll-guard.js deve essere network-first']
];

const failures = checks.filter(([ok]) => !ok).map(([, message]) => message);
if (failures.length) {
  console.error('Controlli stabilità falliti:');
  failures.forEach((failure) => console.error(`- ${failure}`));
  process.exit(1);
}
console.log(`OK: ${checks.length} controlli stabilità Impianti consigliati superati.`);

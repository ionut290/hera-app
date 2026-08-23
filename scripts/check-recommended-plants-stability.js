const fs = require('fs');
const path = require('path');

const root = process.argv[2] || path.resolve(__dirname, '..');
const read = (name) => fs.readFileSync(path.join(root, name), 'utf8');
const core = read('recommended-plants.js');
const css = read('recommended-plants.css');
const adaptive = read('adaptive-work-learning.js');
const equipment = read('equipment-recommendations.js');
const guard = read('fatto-scroll-guard.js');
const squad = read('squad-context-bridge.js');
const traffic = read('recommended-traffic-weather.js');
const streetView = read('street-view-cards.js');
const headers = read('_headers');
const sw = read('sw.js');

const checks = [
  [!adaptive.includes('new MutationObserver'), 'adaptive-work-learning non deve creare observer ricorsivi'],
  [!equipment.includes('new MutationObserver'), 'equipment-recommendations non deve creare observer ricorsivi'],
  [!guard.includes('equipment-recommendations.js'), 'il lazy loader non deve ricaricare il vecchio motore attrezzature'],
  [!guard.includes('adaptive-work-learning.js'), 'il lazy loader non deve ricaricare il vecchio motore adattivo'],
  [guard.includes('scheduleStreetView'), 'Street View deve essere pianificato separatamente'],
  [guard.includes('recommendedPanelOpen()'), 'Street View non deve partire mentre il pannello consigliati è aperto'],
  [/output\.length\s*<\s*CFG\.max/.test(core), 'il percorso deve fermarsi agli impianti visibili della giornata'],
  [core.includes('yieldToInput') && core.includes('maybeYield'), 'il calcolo deve cedere il controllo agli input utente'],
  [core.includes('inputBudgetMs'), 'il lavoro sincrono deve avere un budget temporale breve'],
  [!core.includes('remaining.sort('), 'il calcolo non deve riordinare ripetutamente tutti gli impianti'],
  [core.includes('function closePanel()'), 'deve esistere la chiusura immediata del pannello'],
  [/state\.seq\s*\+=\s*1/.test(core), 'la chiusura deve annullare il calcolo ancora in corso'],
  [core.includes('#back-to-home-btn') && core.includes('pointerdown'), 'il ritorno Home deve annullare il calcolo già al pointerdown'],
  [!core.match(/observe\(page,[\s\S]{0,180}childList\s*:\s*true/), 'il core non deve osservare tutto il sottoalbero della pagina'],
  [/coordinateCache\s*=\s*new WeakMap/.test(core), 'le coordinate devono essere memorizzate per evitare ricalcoli'],
  [/workSourceCache\s*=\s*new WeakMap/.test(core), 'il testo lavorazioni deve essere memorizzato per evitare ricalcoli'],
  [!squad.includes('sync({ force: true })'), 'i tocchi non devono forzare ricalcoli identici della squadra'],
  [squad.includes('applySquadContext'), 'il ponte squadra deve aggiornare solo il contesto cambiato'],
  [!traffic.includes('annotatePanel'), 'traffico/meteo non deve riscrivere direttamente il DOM'],
  [traffic.includes('writeJson(ROUTE_CACHE_KEY, routeCache)'), 'la cache traffico deve essere salvata una sola volta per elaborazione'],
  [traffic.includes("addEventListener('hera:recommended-ready'"), 'traffico/meteo deve partire dopo il rendering del pannello'],
  [streetView.includes('ROWS_PER_BATCH'), 'Street View deve modificare le righe a piccoli gruppi'],
  [streetView.includes('requestIdleCallback(processQueue'), 'Street View deve lavorare nei momenti liberi della UI'],
  [!streetView.includes("querySelectorAll('.impianto-primary-actions').forEach(enhanceRow)"), 'Street View non deve modificare tutte le righe in modo sincrono'],
  [!css.match(/z-index\s*:\s*40/i), 'il pannello non deve usare z-index 40'],
  [!css.match(/isolation\s*:\s*isolate/i), 'il pannello non deve creare uno stacking context isolato'],
  [css.includes('min-height: 44px'), 'i controlli touch devono avere un bersaglio di almeno 44 px'],
  [css.includes('.recommended-plants-panel.hidden'), 'il pannello nascosto deve essere realmente rimosso dai tocchi'],
  [headers.includes('/recommended-plants.js') && headers.includes('/street-view-cards.js'), 'gli asset consigliati devono essere sempre rivalidati dalla rete'],
  [sw.includes('varga-cantieri-shell-v137'), 'il service worker deve mantenere una cache PWA versionata'],
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
console.log(`OK: ${checks.length} controlli reattività Impianti consigliati superati.`);

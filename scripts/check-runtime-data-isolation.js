'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const bridgePath = path.join(root, 'squad-context-bridge.js');
const guardPath = path.join(root, 'fatto-scroll-guard.js');
const bridgeSource = fs.readFileSync(bridgePath, 'utf8');
const guardSource = fs.readFileSync(guardPath, 'utf8');

// Controllo sintattico dei due moduli modificati.
new vm.Script(bridgeSource, { filename: bridgePath });
new vm.Script(guardSource, { filename: guardPath });

// Il modulo delle stime non deve mai scrivere nello stato reale delle squadre.
assert.doesNotMatch(
  bridgeSource,
  /window\.currentSquadre\s*=/,
  'squad-context-bridge.js non deve sovrascrivere window.currentSquadre'
);
assert.match(
  bridgeSource,
  /window\.HeraRecommendedSquadContext\s*=\s*context/,
  'Il contesto delle stime deve usare uno spazio dedicato'
);

// Simula una squadra reale e verifica che il bridge la lasci invariata.
class FixedDate extends Date {
  constructor(...args) {
    super(...(args.length ? args : [2026, 7, 23, 10, 0, 0]));
  }
}

const realSquads = [
  { id: 'real-team-1', commessaId: 'commessa-1' },
  { id: 'real-team-2', commessaId: 'commessa-2' }
];
const originalSquadsJson = JSON.stringify(realSquads);
const history = new Map([
  ['2026-08-23', new Map([
    ['commessa-1', {
      squadre: [{
        caposquadra: 'Mario Rossi',
        personale: 'Luigi Bianchi',
        mezzi: 'A113, T12'
      }]
    }]
  ])]
]);
let intervalCalls = 0;

const documentStub = {
  hidden: false,
  addEventListener() {}
};
const windowStub = {
  selectedCommessaId: 'commessa-1',
  currentSquadre: realSquads,
  addEventListener() {},
  dispatchEvent() {},
  setTimeout() { return 1; },
  clearTimeout() {},
  setInterval() { intervalCalls += 1; return intervalCalls; },
  clearInterval() {}
};
windowStub.window = windowStub;

class CustomEventStub {
  constructor(type, init = {}) {
    this.type = type;
    this.detail = init.detail;
  }
}

const sandbox = {
  window: windowStub,
  document: documentStub,
  CustomEvent: CustomEventStub,
  Date: FixedDate,
  Map,
  Set,
  JSON,
  Math,
  Number,
  String,
  Array,
  Object,
  RegExp,
  console,
  clearTimeout() {},
  selectedCommessaId: 'commessa-1',
  currentSquadre: realSquads,
  squadreHistoryByDate: history,
  doesSquadraMemberMatchCurrentUser(value) {
    return String(value || '').includes('Mario Rossi');
  }
};

vm.runInNewContext(bridgeSource, sandbox, { filename: bridgePath });

assert.strictEqual(windowStub.currentSquadre, realSquads, 'Il riferimento alle squadre reali è cambiato');
assert.equal(JSON.stringify(windowStub.currentSquadre), originalSquadsJson, 'Il contenuto delle squadre reali è cambiato');
assert.equal(windowStub.HeraRecommendedSquadContext?.commessaId, 'commessa-1');
assert.equal(windowStub.HeraRecommendedSquadContext?.teamSize, 2);
assert.equal(windowStub.HeraRecommendedSquadContext?.hasDaily, true);
assert.equal(windowStub.HeraRecommendedSquadContext?.hasTrincia, true);
assert.equal(intervalCalls, 0, 'Il polling non deve partire quando il contesto è già disponibile');

// I moduli accessori devono partire solo nella pagina Impianti o su azione esplicita.
assert.match(guardSource, /function\s+initializeLazyHelpers\s*\(/);
assert.match(guardSource, /getElementById\(['"]impianti-page['"]\)/);
assert.match(guardSource, /requestIdleCallback/);
assert.match(guardSource, /20260823-stability1/);
assert.doesNotMatch(
  guardSource,
  /DOMContentLoaded['"],\s*loadRecommendedHelpers/,
  'I moduli accessori non devono essere caricati direttamente al DOMContentLoaded'
);
assert.doesNotMatch(
  guardSource,
  /else\s*{\s*loadRecommendedHelpers\(\);\s*}\s*\)\(\);?\s*$/,
  'I moduli accessori non devono essere caricati subito quando il DOM è già pronto'
);

console.log('OK: stato squadre isolato e moduli Impianti caricati in modo differito.');

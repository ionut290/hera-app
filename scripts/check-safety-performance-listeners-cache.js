const fs = require('fs');
const path = require('path');

const root = process.cwd();
const fullPath = (file) => path.join(root, file);
const read = (file) => fs.readFileSync(fullPath(file), 'utf8');
const size = (file) => fs.statSync(fullPath(file)).size;
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
  console.log(`✓ ${message}`);
};

function findExplicitWhileTrueBodies(source) {
  const bodies = [];
  const loopPattern = /\bwhile\s*\(\s*true\s*\)\s*\{/g;
  let match;

  while ((match = loopPattern.exec(source)) !== null) {
    const bodyStart = loopPattern.lastIndex;
    let depth = 1;
    let index = bodyStart;
    let quote = null;
    let escaped = false;
    let lineComment = false;
    let blockComment = false;

    for (; index < source.length && depth > 0; index += 1) {
      const char = source[index];
      const next = source[index + 1];

      if (lineComment) {
        if (char === '\n') lineComment = false;
        continue;
      }

      if (blockComment) {
        if (char === '*' && next === '/') {
          blockComment = false;
          index += 1;
        }
        continue;
      }

      if (quote) {
        if (escaped) {
          escaped = false;
          continue;
        }
        if (char === '\\') {
          escaped = true;
          continue;
        }
        if (char === quote) quote = null;
        continue;
      }

      if (char === '/' && next === '/') {
        lineComment = true;
        index += 1;
        continue;
      }

      if (char === '/' && next === '*') {
        blockComment = true;
        index += 1;
        continue;
      }

      if (char === '"' || char === "'" || char === '`') {
        quote = char;
        continue;
      }

      if (char === '{') depth += 1;
      if (char === '}') depth -= 1;
    }

    const bodyEnd = depth === 0 ? index - 1 : source.length;
    bodies.push(source.slice(bodyStart, bodyEnd));
    loopPattern.lastIndex = Math.max(loopPattern.lastIndex, index);
  }

  return bodies;
}

const required = [
  'index.html',
  'app.js',
  'sw.js',
  'firestore.indexes.json',
  'firestore-safe-optimizer.js',
  'firestore-inflight-read-coalescer.js',
  'firestore-diagnostics-optimizer-extension.js',
  'scripts/check-firestore-safe-optimizer.js',
  'scripts/check-firestore-inflight-read-coalescer.js'
];

for (const file of required) {
  assert(fs.existsSync(fullPath(file)) && size(file) > 0, `${file} esiste e non è vuoto`);
}

const index = read('index.html');
const app = read('app.js');
const sw = read('sw.js');
const indexes = JSON.parse(read('firestore.indexes.json'));
const safeOptimizer = read('firestore-safe-optimizer.js');
const inflightOptimizer = read('firestore-inflight-read-coalescer.js');
const diagnosticsOptimizer = read('firestore-diagnostics-optimizer-extension.js');

// Budget molto larghi: segnalano soltanto aumenti anomali, non normali sviluppi dell'app.
assert(size('index.html') < 2 * 1024 * 1024, 'index.html resta sotto 2 MB');
assert(size('app.js') < 6 * 1024 * 1024, 'app.js resta sotto 6 MB');
assert(size('sw.js') < 1024 * 1024, 'sw.js resta sotto 1 MB');

// Cache PWA: nome versionato, pulizia vecchie cache e attivazione controllata.
assert(/CACHE_NAME\s*=\s*['"][^'"]*v\d+[^'"]*['"]/.test(sw), 'La cache PWA usa un nome versionato');
assert(/caches\.keys\s*\(/.test(sw) && /caches\.delete\s*\(/.test(sw), 'Il Service Worker elimina le vecchie cache');
assert(/skipWaiting\s*\(/.test(sw), 'Il Service Worker attiva rapidamente la nuova versione');
assert(/clients\.claim\s*\(/.test(sw), 'Il Service Worker prende il controllo delle pagine aperte');
assert(/addEventListener\s*\(\s*['"]fetch['"]/.test(sw), 'La strategia cache gestisce le richieste fetch');

const appVersionIndex = index.match(/app\.js\?v=([^"']+)/)?.[1];
const appVersionSw = sw.match(/app\.js\?v=([^"']+)/)?.[1];
assert(Boolean(appVersionIndex), 'index.html usa una versione esplicita di app.js');
assert(appVersionIndex === appVersionSw, 'La cache usa la stessa versione di app.js caricata dalla pagina');

// Protezioni contro listener e letture duplicate.
assert(/onSnapshot|snapshot/i.test(safeOptimizer), 'L’ottimizzatore listener gestisce gli snapshot Firestore');
assert(/unsubscribe|detach|release|subscribers|listeners/i.test(safeOptimizer), 'L’ottimizzatore conserva una gestione del ciclo di vita dei listener');
assert(/inflight|pending|Map\s*\(|promise/i.test(inflightOptimizer), 'Le letture simultanee identiche vengono aggregate');
assert(/diagnostic|metric|operation|read/i.test(diagnosticsOptimizer), 'La diagnostica continua a osservare le operazioni Firestore');
assert(/firestore-safe-optimizer\.js/.test(index + sw), 'L’ottimizzatore listener è collegato al bootstrap o alla cache');
assert(/firestore-inflight-read-coalescer\.js/.test(index + sw), 'L’ottimizzatore delle letture simultanee è collegato al bootstrap o alla cache');

// Indici Firestore: struttura valida e nessun indice vuoto o duplicato.
assert(Array.isArray(indexes.indexes), 'firestore.indexes.json contiene un elenco di indici');
assert(Array.isArray(indexes.fieldOverrides), 'firestore.indexes.json contiene fieldOverrides');
const signatures = new Set();
for (const [position, entry] of indexes.indexes.entries()) {
  assert(typeof entry.collectionGroup === 'string' && entry.collectionGroup.trim(), `Indice ${position + 1}: collectionGroup valido`);
  assert(['COLLECTION', 'COLLECTION_GROUP'].includes(entry.queryScope), `Indice ${position + 1}: queryScope valido`);
  assert(Array.isArray(entry.fields) && entry.fields.length >= 2, `Indice ${position + 1}: almeno due campi`);
  for (const field of entry.fields) {
    assert(typeof field.fieldPath === 'string' && field.fieldPath.trim(), `Indice ${position + 1}: fieldPath valido`);
    assert(Boolean(field.order || field.arrayConfig), `Indice ${position + 1}: ordinamento o arrayConfig presente`);
  }
  const signature = JSON.stringify(entry);
  assert(!signatures.has(signature), `Indice ${position + 1}: nessun duplicato identico`);
  signatures.add(signature);
}

// Evita sorgenti e mappe di debug enormi nel caricamento iniziale.
assert(!/<script[^>]+src=["'][^"']+\.map(?:\?|["'])/i.test(index), 'La pagina non carica source map come script');

const whileTrueBodies = findExplicitWhileTrueBodies(app);
const suspiciousWhileTrueBodies = whileTrueBodies.filter(
  (body) => !/\b(?:break|return|throw)\b/.test(body)
);
assert(
  suspiciousWhileTrueBodies.length === 0,
  'Ogni while (true) in app.js contiene un’uscita esplicita'
);

console.log('\nControlli prestazioni, listener, cache PWA e indici completati senza connessioni esterne.');

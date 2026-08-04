(() => {
  'use strict';
  if (window.__vargaFsDiagV2) return;
  window.__vargaFsDiagV2 = true;

  const PREFIX = 'varga_fs_diag_v2_';
  const batchOps = new WeakMap();
  let patched = false;
  let renderTimer = 0;
  let activeListeners = 0;

  const text = (value) => String(value ?? '').trim();
  const today = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Europe/Rome' });
  const storageKey = () => PREFIX + today();
  const blank = () => ({
    version: 2,
    date: today(),
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    reads: 0,
    writes: 0,
    deletes: 0,
    listenerRegistrations: 0,
    listenerUnsubscribes: 0,
    activeListeners: 0,
    peakActiveListeners: 0,
    areas: {},
    collections: {},
    methods: {},
    callers: {},
    details: []
  });

  function load() {
    try {
      const value = JSON.parse(localStorage.getItem(storageKey()) || 'null');
      return value?.date === today() ? { ...blank(), ...value } : blank();
    } catch (_) {
      return blank();
    }
  }

  function save(value) {
    try {
      value.updatedAt = new Date().toISOString();
      value.activeListeners = activeListeners;
      value.peakActiveListeners = Math.max(Number(value.peakActiveListeners || 0), activeListeners);
      localStorage.setItem(storageKey(), JSON.stringify(value));
    } catch (_) {}
    clearTimeout(renderTimer);
    renderTimer = setTimeout(render, 150);
  }

  function increment(object, key, amount) {
    object[key] = (Number(object[key]) || 0) + amount;
  }

  function firstCollection(path) {
    const value = text(path);
    const match = value.match(/(?:^|\|path=)([^/|]+)/);
    return match?.[1] || value.split('/').filter(Boolean)[0] || 'sconosciuta';
  }

  function classify(path, caller = '') {
    const value = `${path} ${caller}`.toLowerCase();
    const rules = [
      ['Preventivi/Consuntivi', /preventiv|consuntiv|quote/],
      ['Contabilità', /accounting|contabil|lavorazion|prezziario|impiantifisici/],
      ['Impianti', /impianti|commesse/],
      ['Ore', /hours|ore|timbr|rapport/],
      ['Squadre', /squadre|team|assegnazion/],
      ['Notifiche', /notification|notific|avvis/],
      ['Personale', /personale|operator|utenti|users/],
      ['Documenti', /document|pos|allegat/],
      ['Segnalazioni', /segnalazion|report/],
      ['Presenza/Posizione', /presence|presenz|position|posizion|location/],
      ['Global', /global/]
    ];
    return rules.find((item) => item[1].test(value))?.[0] || 'Altro';
  }

  function callerFromStack(stack) {
    const lines = text(stack).split('\n').map((line) => line.trim());
    const useful = lines.find((line) =>
      line &&
      !/firestore-operation-diagnostics\.js/i.test(line) &&
      !/firebase(?:-firestore)?(?:\.js)?/i.test(line) &&
      !/new Error/i.test(line)
    );
    return useful || 'chiamante non identificato';
  }

  function queryDescription(query) {
    const internal = query?._query || query?.Ae || query?.je || null;
    const path =
      internal?.path?.canonicalString?.() ||
      internal?.path?.toString?.() ||
      internal?.path?.segments?.join?.('/') ||
      query?.path ||
      query?._delegate?._query?.path?.canonicalString?.() ||
      'query-senza-percorso';
    let canonical = '';
    try {
      canonical = internal?.canonicalId?.() || internal?.toString?.() || '';
    } catch (_) {}
    return canonical && canonical !== '[object Object]' ? `${path}|query=${canonical}` : path;
  }

  function record(type, method, path, amount = 1, stack = '') {
    amount = Math.max(0, Number(amount) || 0);
    if (!amount) return;
    const caller = callerFromStack(stack);
    const value = load();
    const area = classify(path, caller);
    if (type === 'read') value.reads += amount;
    if (type === 'write') value.writes += amount;
    if (type === 'delete') value.deletes += amount;
    if (type === 'listener-open') value.listenerRegistrations += amount;
    if (type === 'listener-close') value.listenerUnsubscribes += amount;
    increment(value.areas, `${area}:${type}`, amount);
    increment(value.collections, `${firstCollection(path)}:${type}`, amount);
    increment(value.methods, `${method}:${type}`, amount);
    increment(value.callers, `${caller}:${type}`, amount);
    value.details.unshift({
      at: new Date().toISOString(),
      type,
      method,
      path: text(path) || 'sconosciuto',
      area,
      caller,
      amount,
      activeListeners
    });
    value.details = value.details.slice(0, 400);
    save(value);
  }

  const quantity = (snapshot) => typeof snapshot?.size === 'number'
    ? snapshot.size
    : typeof snapshot?.exists === 'boolean'
      ? (snapshot.exists ? 1 : 0)
      : 1;

  function observePromise(result, callback) {
    try {
      if (result?.then) result.then(callback).catch(() => {});
    } catch (_) {}
    return result;
  }

  function wrap(prototype, name, factory) {
    if (!prototype || typeof prototype[name] !== 'function' || prototype[name].__vargaDiagV2) return;
    const original = prototype[name];
    const wrapped = factory(original);
    wrapped.__vargaDiagV2 = true;
    wrapped.__vargaOriginal = original;
    prototype[name] = wrapped;
  }

  function listenerWrapper(original, method, pathFactory) {
    return function (...args) {
      const stack = new Error().stack;
      const path = pathFactory(this);
      activeListeners += 1;
      record('listener-open', method, path, 1, stack);
      let unsubscribe;
      try {
        unsubscribe = original.apply(this, args);
      } catch (error) {
        activeListeners = Math.max(0, activeListeners - 1);
        record('listener-close', `${method}.failed`, path, 1, stack);
        throw error;
      }
      if (typeof unsubscribe !== 'function') return unsubscribe;
      let closed = false;
      return function diagnosticUnsubscribe() {
        if (!closed) {
          closed = true;
          activeListeners = Math.max(0, activeListeners - 1);
          record('listener-close', `${method}.unsubscribe`, path, 1, new Error().stack);
        }
        return unsubscribe.apply(this, arguments);
      };
    };
  }

  function patch() {
    if (patched || !window.firebase?.firestore) return false;
    const firestore = window.firebase.firestore;
    const DocumentReference = firestore.DocumentReference?.prototype;
    const CollectionReference = firestore.CollectionReference?.prototype;
    const Query = firestore.Query?.prototype;
    const WriteBatch = firestore.WriteBatch?.prototype;
    const Transaction = firestore.Transaction?.prototype;

    wrap(DocumentReference, 'set', (original) => function (...args) {
      const stack = new Error().stack;
      return observePromise(original.apply(this, args), () => record('write', 'doc.set', this.path, 1, stack));
    });
    wrap(DocumentReference, 'update', (original) => function (...args) {
      const stack = new Error().stack;
      return observePromise(original.apply(this, args), () => record('write', 'doc.update', this.path, 1, stack));
    });
    wrap(DocumentReference, 'delete', (original) => function (...args) {
      const stack = new Error().stack;
      return observePromise(original.apply(this, args), () => record('delete', 'doc.delete', this.path, 1, stack));
    });
    wrap(DocumentReference, 'get', (original) => function (...args) {
      const stack = new Error().stack;
      return observePromise(original.apply(this, args), (snapshot) => record('read', 'doc.get', this.path, quantity(snapshot), stack));
    });
    wrap(DocumentReference, 'onSnapshot', (original) => listenerWrapper(original, 'doc.onSnapshot', (reference) => reference.path));

    wrap(CollectionReference, 'add', (original) => function (...args) {
      const stack = new Error().stack;
      return observePromise(original.apply(this, args), () => record('write', 'collection.add', this.path, 1, stack));
    });

    wrap(Query, 'get', (original) => function (...args) {
      const stack = new Error().stack;
      const path = queryDescription(this);
      return observePromise(original.apply(this, args), (snapshot) => record('read', 'query.get', path, quantity(snapshot), stack));
    });
    wrap(Query, 'onSnapshot', (original) => listenerWrapper(original, 'query.onSnapshot', queryDescription));

    ['set', 'update', 'delete'].forEach((method) => wrap(WriteBatch, method, (original) => function (reference, ...args) {
      const operations = batchOps.get(this) || [];
      operations.push({ method, path: reference?.path || 'batch', stack: new Error().stack });
      batchOps.set(this, operations);
      return original.call(this, reference, ...args);
    }));
    wrap(WriteBatch, 'commit', (original) => function (...args) {
      const operations = (batchOps.get(this) || []).slice();
      const result = original.apply(this, args);
      observePromise(result, () => {
        operations.forEach((entry) => record(entry.method === 'delete' ? 'delete' : 'write', `batch.${entry.method}`, entry.path, 1, entry.stack));
        batchOps.delete(this);
      });
      return result;
    });

    ['set', 'update', 'delete'].forEach((method) => wrap(Transaction, method, (original) => function (reference, ...args) {
      record(method === 'delete' ? 'delete' : 'write', `transaction.${method}`, reference?.path || 'transaction', 1, new Error().stack);
      return original.call(this, reference, ...args);
    }));
    wrap(Transaction, 'get', (original) => function (reference, ...args) {
      const stack = new Error().stack;
      return observePromise(original.call(this, reference, ...args), (snapshot) => record('read', 'transaction.get', reference?.path || 'transaction', quantity(snapshot), stack));
    });

    patched = true;
    return true;
  }

  const escapeHtml = (value) => text(value).replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[character]));

  function rows(map, suffix) {
    return Object.entries(map || {})
      .filter(([key]) => key.endsWith(`:${suffix}`))
      .map(([key, value]) => [key.slice(0, -suffix.length - 1), value])
      .sort((a, b) => b[1] - a[1]);
  }

  function ensureCard() {
    let card = document.getElementById('firestore-operation-diagnostics-card');
    if (card) return card;
    const root = document.getElementById('control-center-content') || document.getElementById('control-center-page');
    if (!root) return null;
    card = document.createElement('section');
    card.id = 'firestore-operation-diagnostics-card';
    card.className = 'card';
    card.innerHTML = '<div class="section-head"><div><h2>🔎 Diagnostica operazioni Firestore</h2><p class="muted">Mostra query, chiamante e listener realmente attivi. Non genera operazioni Firestore.</p></div><button class="btn" data-refresh>AGGIORNA</button></div><div data-summary></div><div class="actions-row"><button class="btn" data-export>SCARICA REPORT</button><button class="btn" data-reset>AZZERA OGGI</button></div>';
    root.appendChild(card);
    card.querySelector('[data-refresh]').onclick = render;
    card.querySelector('[data-reset]').onclick = () => {
      if (confirm('Azzerare la diagnostica locale e iniziare una nuova sessione?')) {
        localStorage.setItem(storageKey(), JSON.stringify(blank()));
        window.location.reload();
      }
    };
    card.querySelector('[data-export]').onclick = () => {
      const value = load();
      const url = URL.createObjectURL(new Blob([JSON.stringify(value, null, 2)], { type: 'application/json' }));
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `diagnostica-firestore-${value.date}.json`;
      anchor.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    };
    return card;
  }

  function render() {
    const card = ensureCard();
    if (!card) return;
    const value = load();
    value.activeListeners = activeListeners;
    const writes = rows(value.areas, 'write');
    const reads = rows(value.areas, 'read');
    const callers = rows(value.callers, 'listener-open').slice(0, 12);
    const collections = Object.entries(value.collections || {}).sort((a, b) => b[1] - a[1]).slice(0, 15);
    card.querySelector('[data-summary]').innerHTML = `
      <div class="control-center-grid">
        <div class="control-center-row"><span>Letture osservate</span><strong>${value.reads}</strong></div>
        <div class="control-center-row"><span>Scritture osservate</span><strong>${value.writes}</strong></div>
        <div class="control-center-row"><span>Eliminazioni osservate</span><strong>${value.deletes}</strong></div>
        <div class="control-center-row"><span>Listener aperti complessivamente</span><strong>${value.listenerRegistrations}</strong></div>
        <div class="control-center-row"><span>Listener chiusi</span><strong>${value.listenerUnsubscribes}</strong></div>
        <div class="control-center-row"><span>Listener attivi adesso</span><strong>${activeListeners}</strong></div>
        <div class="control-center-row"><span>Picco listener attivi</span><strong>${Math.max(value.peakActiveListeners || 0, activeListeners)}</strong></div>
      </div>
      <details open><summary><strong>Chiamanti dei listener</strong></summary>${callers.length ? callers.map(([name, count]) => `<div class="control-center-row"><span>${escapeHtml(name)}</span><strong>${count}</strong></div>`).join('') : '<p class="muted">Nessun listener osservato.</p>'}</details>
      <details><summary><strong>Operazioni per percorso/query</strong></summary>${collections.length ? collections.map(([name, count]) => `<div class="control-center-row"><span>${escapeHtml(name.replace(':', ' · '))}</span><strong>${count}</strong></div>`).join('') : '<p class="muted">Nessuna operazione registrata.</p>'}</details>
      <details><summary><strong>Scritture per area</strong></summary>${writes.length ? writes.map(([name, count]) => `<div class="control-center-row"><span>${escapeHtml(name)}</span><strong>${count}</strong></div>`).join('') : '<p class="muted">Nessuna scrittura osservata.</p>'}</details>
      <details><summary><strong>Letture per area</strong></summary>${reads.length ? reads.map(([name, count]) => `<div class="control-center-row"><span>${escapeHtml(name)}</span><strong>${count}</strong></div>`).join('') : '<p class="muted">Nessuna lettura osservata.</p>'}</details>
      <p class="muted">Monitor V2 attivo dal ${new Date(value.startedAt).toLocaleString('it-IT')}. Il report JSON include percorso completo, query, chiamante e numero di listener attivi.</p>`;
  }

  function init() {
    if (!patch()) {
      let attempts = 0;
      const interval = setInterval(() => {
        attempts += 1;
        if (patch() || attempts >= 40) clearInterval(interval);
      }, 250);
    }
    render();
    const root = document.getElementById('control-center-content') || document.getElementById('control-center-page');
    if (root) new MutationObserver(() => {
      if (!document.getElementById('firestore-operation-diagnostics-card')) setTimeout(render, 0);
    }).observe(root, { childList: true, subtree: false });
  }

  window.VargaFirestoreDiagnostics = {
    read: load,
    render,
    reset: () => { localStorage.setItem(storageKey(), JSON.stringify(blank())); render(); },
    activeListeners: () => activeListeners
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();

(() => {
  'use strict';
  if (window.__vargaFsDiagV3) return;
  window.__vargaFsDiagV3 = true;

  const PREFIX_V3 = 'varga_fs_diag_v3_';
  const PREFIX_V2 = 'varga_fs_diag_v2_';
  const SCRIPT_VERSION = '3.1.0';
  const DETAIL_LIMIT = 2000;
  const batchOps = new WeakMap();
  const liveListeners = new Map();
  let patched = false;
  let renderTimer = 0;
  let activeListeners = 0;
  let listenerSequence = 0;

  const text = (value) => String(value ?? '').trim();
  const nowIso = () => new Date().toISOString();
  const today = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Europe/Rome' });
  const storageKey = () => PREFIX_V3 + today();
  const legacyStorageKey = () => PREFIX_V2 + today();
  const sessionId = () => `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
  const numeric = (value) => Math.max(0, Number(value) || 0);

  const blank = () => ({
    version: 3,
    scriptVersion: SCRIPT_VERSION,
    date: today(),
    sessionId: sessionId(),
    startedAt: nowIso(),
    updatedAt: nowIso(),
    initialUrl: typeof location !== 'undefined' ? location.href : '',
    userAgent: typeof navigator !== 'undefined' ? navigator.userAgent : '',
    buildVersion: window.APP_VERSION || window.BUILD_VERSION || '',
    reads: 0,
    readOperations: 0,
    readDocuments: 0,
    readLatencyMsTotal: 0,
    readLatencyMsMax: 0,
    listenerDeliveries: 0,
    listenerDocuments: 0,
    unattributedReads: 0,
    writes: 0,
    deletes: 0,
    listenerRegistrations: 0,
    listenerUnsubscribes: 0,
    activeListeners: 0,
    peakActiveListeners: 0,
    detailsDropped: 0,
    areas: {},
    collections: {},
    methods: {},
    callers: {},
    queries: {},
    functions: {},
    screens: {},
    listenerInstances: {},
    details: []
  });

  function migrate(value) {
    const base = blank();
    if (!value || value.date !== today()) return base;
    const migrated = { ...base, ...value, version: 3, scriptVersion: SCRIPT_VERSION };
    migrated.readDocuments = numeric(value.readDocuments ?? value.reads);
    migrated.reads = migrated.readDocuments;
    migrated.readOperations = numeric(value.readOperations);
    migrated.readLatencyMsTotal = numeric(value.readLatencyMsTotal);
    migrated.readLatencyMsMax = numeric(value.readLatencyMsMax);
    migrated.listenerDeliveries = numeric(value.listenerDeliveries);
    migrated.listenerDocuments = numeric(value.listenerDocuments);
    migrated.unattributedReads = numeric(value.unattributedReads);
    migrated.detailsDropped = numeric(value.detailsDropped);
    ['areas', 'collections', 'methods', 'callers', 'queries', 'functions', 'screens', 'listenerInstances'].forEach((key) => {
      migrated[key] = value[key] && typeof value[key] === 'object' ? value[key] : {};
    });
    migrated.details = Array.isArray(value.details) ? value.details.slice(0, DETAIL_LIMIT) : [];
    return migrated;
  }

  function load() {
    try {
      const current = JSON.parse(localStorage.getItem(storageKey()) || 'null');
      if (current?.date === today()) return migrate(current);
      const legacy = JSON.parse(localStorage.getItem(legacyStorageKey()) || 'null');
      if (legacy?.date === today()) return migrate(legacy);
    } catch (_) {}
    return blank();
  }

  function safePersist(value) {
    value.updatedAt = nowIso();
    value.activeListeners = activeListeners;
    value.peakActiveListeners = Math.max(numeric(value.peakActiveListeners), activeListeners);
    let details = Array.isArray(value.details) ? value.details : [];
    while (true) {
      try {
        value.details = details;
        localStorage.setItem(storageKey(), JSON.stringify(value));
        break;
      } catch (_) {
        if (!details.length) break;
        const remove = Math.max(1, Math.ceil(details.length / 4));
        details = details.slice(0, Math.max(0, details.length - remove));
        value.detailsDropped = numeric(value.detailsDropped) + remove;
      }
    }
    clearTimeout(renderTimer);
    renderTimer = setTimeout(render, 150);
  }

  function increment(object, key, amount) {
    object[key] = numeric(object[key]) + numeric(amount);
  }

  function incrementStats(object, key, docs, operations = 0, durationMs = 0) {
    const current = object[key] && typeof object[key] === 'object' ? object[key] : {};
    current.documents = numeric(current.documents) + numeric(docs);
    current.operations = numeric(current.operations) + numeric(operations);
    current.durationMsTotal = numeric(current.durationMsTotal) + numeric(durationMs);
    current.durationMsMax = Math.max(numeric(current.durationMsMax), numeric(durationMs));
    object[key] = current;
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
    return lines.find((line) => line &&
      !/firestore-operation-diagnostics\.js/i.test(line) &&
      !/firebase(?:-firestore)?(?:\.js)?/i.test(line) &&
      !/new Error/i.test(line)) || 'chiamante non identificato';
  }

  function functionNameFromCaller(caller) {
    const value = text(caller);
    const match = value.match(/(?:at\s+)?([^\s@()]+)\s*(?:@|\()/);
    const name = text(match?.[1]);
    if (!name || /https?:|anonymous|<anonymous>|chiamante/i.test(name)) return 'non attribuita';
    return name.replace(/^Object\./, '').replace(/^window\./, '');
  }

  function currentScreen() {
    try {
      const visible = Array.from(document.querySelectorAll('[data-page], .page, section[id], main[id]'))
        .find((element) => element.offsetParent !== null && getComputedStyle(element).display !== 'none');
      const label = visible?.dataset?.page || visible?.id;
      if (label) return label;
      return location.hash || location.pathname || 'sconosciuta';
    } catch (_) {
      return 'sconosciuta';
    }
  }

  function queryDescription(query) {
    const internal = query?._query || query?.Ae || query?.je || null;
    const path = internal?.path?.canonicalString?.() || internal?.path?.toString?.() ||
      internal?.path?.segments?.join?.('/') || query?.path ||
      query?._delegate?._query?.path?.canonicalString?.() || 'query-senza-percorso';
    let canonical = '';
    try { canonical = internal?.canonicalId?.() || internal?.toString?.() || ''; } catch (_) {}
    return canonical && canonical !== '[object Object]' ? `${path}|query=${canonical}` : path;
  }

  const quantity = (snapshot) => typeof snapshot?.size === 'number'
    ? snapshot.size
    : typeof snapshot?.exists === 'boolean'
      ? (snapshot.exists ? 1 : 0)
      : 1;

  function listenerQuantity(snapshot, deliveryNumber) {
    if (deliveryNumber <= 1 || typeof snapshot?.docChanges !== 'function') return quantity(snapshot);
    try {
      return snapshot.docChanges().length;
    } catch (_) {
      return quantity(snapshot);
    }
  }

  function addDetail(value, detail) {
    value.details.unshift(detail);
    if (value.details.length > DETAIL_LIMIT) {
      value.detailsDropped += value.details.length - DETAIL_LIMIT;
      value.details.length = DETAIL_LIMIT;
    }
  }

  function record(type, method, path, amount = 1, stack = '', extra = {}) {
    amount = numeric(amount);
    if (!amount && !['listener-open', 'listener-close', 'listener-delivery'].includes(type)) return;
    const caller = extra.caller || callerFromStack(stack);
    const functionName = extra.functionName || functionNameFromCaller(caller);
    const screen = extra.screen || currentScreen();
    const collection = firstCollection(path);
    const area = classify(path, caller);
    const durationMs = numeric(extra.durationMs);
    const value = load();

    if (type === 'read') {
      value.readOperations += 1;
      value.readDocuments += amount;
      value.reads = value.readDocuments;
      value.readLatencyMsTotal += durationMs;
      value.readLatencyMsMax = Math.max(value.readLatencyMsMax, durationMs);
      if (functionName === 'non attribuita') value.unattributedReads += amount;
    }
    if (type === 'listener-delivery') {
      value.listenerDeliveries += 1;
      value.listenerDocuments += amount;
      value.readDocuments += amount;
      value.reads = value.readDocuments;
      if (functionName === 'non attribuita') value.unattributedReads += amount;
    }
    if (type === 'write') value.writes += amount;
    if (type === 'delete') value.deletes += amount;
    if (type === 'listener-open') value.listenerRegistrations += 1;
    if (type === 'listener-close') value.listenerUnsubscribes += 1;

    increment(value.areas, `${area}:${type}`, amount || 1);
    increment(value.collections, `${collection}:${type}`, amount || 1);
    increment(value.methods, `${method}:${type}`, amount || 1);
    increment(value.callers, `${caller}:${type}`, amount || 1);

    if (type === 'read' || type === 'listener-delivery') {
      incrementStats(value.queries, `${method}|${path}`, amount, 1, durationMs);
      incrementStats(value.functions, functionName, amount, 1, durationMs);
      incrementStats(value.screens, screen, amount, 1, durationMs);
    }

    if (extra.listenerId) {
      const listener = value.listenerInstances[extra.listenerId] || {};
      Object.assign(listener, {
        id: extra.listenerId,
        method: listener.method || method,
        path: listener.path || path,
        collection,
        area,
        caller,
        functionName,
        screen,
        openedAt: listener.openedAt || extra.openedAt || nowIso(),
        closedAt: extra.closedAt || listener.closedAt || null,
        durationMs: extra.listenerDurationMs ?? listener.durationMs ?? null,
        active: extra.active ?? listener.active ?? true,
        deliveries: numeric(listener.deliveries) + (type === 'listener-delivery' ? 1 : 0),
        documents: numeric(listener.documents) + (type === 'listener-delivery' ? amount : 0)
      });
      value.listenerInstances[extra.listenerId] = listener;
    }

    addDetail(value, {
      at: extra.at || nowIso(),
      finishedAt: extra.finishedAt || null,
      durationMs,
      type,
      method,
      path: text(path) || 'sconosciuto',
      collection,
      area,
      caller,
      functionName,
      screen,
      amount,
      listenerId: extra.listenerId || null,
      deliveryNumber: extra.deliveryNumber || null,
      initialDelivery: Boolean(extra.initialDelivery),
      activeListeners
    });
    safePersist(value);
  }

  function observePromise(result, callback) {
    try { if (result?.then) result.then(callback).catch(() => {}); } catch (_) {}
    return result;
  }

  function wrap(prototype, name, factory) {
    if (!prototype || typeof prototype[name] !== 'function' || prototype[name].__vargaDiagV3) return;
    const original = prototype[name];
    const wrapped = factory(original);
    wrapped.__vargaDiagV3 = true;
    wrapped.__vargaOriginal = original;
    prototype[name] = wrapped;
  }

  function wrapSnapshotArgs(args, onDelivery) {
    const wrapped = args.slice();
    const observerIndex = wrapped.findIndex((arg) => arg && typeof arg === 'object' && typeof arg.next === 'function');
    if (observerIndex >= 0) {
      const observer = wrapped[observerIndex];
      wrapped[observerIndex] = { ...observer, next(snapshot) { onDelivery(snapshot); return observer.next.call(observer, snapshot); } };
      return wrapped;
    }
    const callbackIndex = wrapped.findIndex((arg) => typeof arg === 'function');
    if (callbackIndex >= 0) {
      const next = wrapped[callbackIndex];
      wrapped[callbackIndex] = function diagnosticNext(snapshot) { onDelivery(snapshot); return next.apply(this, arguments); };
    }
    return wrapped;
  }

  function listenerWrapper(original, method, pathFactory) {
    return function (...args) {
      const stack = new Error().stack;
      const caller = callerFromStack(stack);
      const functionName = functionNameFromCaller(caller);
      const screen = currentScreen();
      const path = pathFactory(this);
      const listenerId = `L${Date.now().toString(36)}-${++listenerSequence}`;
      const openedAtMs = Date.now();
      const openedAt = new Date(openedAtMs).toISOString();
      let deliveryNumber = 0;
      activeListeners += 1;
      liveListeners.set(listenerId, { listenerId, method, path, caller, functionName, screen, openedAtMs });
      record('listener-open', method, path, 1, stack, { listenerId, caller, functionName, screen, openedAt, active: true });

      const wrappedArgs = wrapSnapshotArgs(args, (snapshot) => {
        deliveryNumber += 1;
        const durationMs = deliveryNumber === 1 ? Date.now() - openedAtMs : 0;
        record('listener-delivery', `${method}.delivery`, path, listenerQuantity(snapshot, deliveryNumber), stack, {
          listenerId, caller, functionName, screen, durationMs,
          deliveryNumber, initialDelivery: deliveryNumber === 1, openedAt, active: true
        });
      });

      let unsubscribe;
      try { unsubscribe = original.apply(this, wrappedArgs); }
      catch (error) {
        activeListeners = Math.max(0, activeListeners - 1);
        liveListeners.delete(listenerId);
        record('listener-close', `${method}.failed`, path, 1, stack, {
          listenerId, caller, functionName, screen, openedAt, closedAt: nowIso(), active: false,
          listenerDurationMs: Date.now() - openedAtMs
        });
        throw error;
      }
      if (typeof unsubscribe !== 'function') return unsubscribe;
      let closed = false;
      return function diagnosticUnsubscribe() {
        if (!closed) {
          closed = true;
          activeListeners = Math.max(0, activeListeners - 1);
          liveListeners.delete(listenerId);
          record('listener-close', `${method}.unsubscribe`, path, 1, new Error().stack, {
            listenerId, caller, functionName, screen, openedAt, closedAt: nowIso(), active: false,
            listenerDurationMs: Date.now() - openedAtMs
          });
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

    ['set', 'update'].forEach((name) => wrap(DocumentReference, name, (original) => function (...args) {
      const stack = new Error().stack;
      return observePromise(original.apply(this, args), () => record('write', `doc.${name}`, this.path, 1, stack));
    }));
    wrap(DocumentReference, 'delete', (original) => function (...args) {
      const stack = new Error().stack;
      return observePromise(original.apply(this, args), () => record('delete', 'doc.delete', this.path, 1, stack));
    });
    wrap(DocumentReference, 'get', (original) => function (...args) {
      const stack = new Error().stack;
      const started = Date.now();
      return observePromise(original.apply(this, args), (snapshot) => record('read', 'doc.get', this.path, quantity(snapshot), stack, {
        durationMs: Date.now() - started, finishedAt: nowIso()
      }));
    });
    wrap(DocumentReference, 'onSnapshot', (original) => listenerWrapper(original, 'doc.onSnapshot', (reference) => reference.path));

    wrap(CollectionReference, 'add', (original) => function (...args) {
      const stack = new Error().stack;
      return observePromise(original.apply(this, args), () => record('write', 'collection.add', this.path, 1, stack));
    });

    wrap(Query, 'get', (original) => function (...args) {
      const stack = new Error().stack;
      const path = queryDescription(this);
      const started = Date.now();
      return observePromise(original.apply(this, args), (snapshot) => record('read', 'query.get', path, quantity(snapshot), stack, {
        durationMs: Date.now() - started, finishedAt: nowIso()
      }));
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
      const started = Date.now();
      return observePromise(original.call(this, reference, ...args), (snapshot) => record('read', 'transaction.get', reference?.path || 'transaction', quantity(snapshot), stack, {
        durationMs: Date.now() - started, finishedAt: nowIso()
      }));
    });

    patched = true;
    return true;
  }

  const escapeHtml = (value) => text(value).replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[character]));
  const sortedStats = (map) => Object.entries(map || {}).sort((a, b) => numeric(b[1]?.documents) - numeric(a[1]?.documents));
  const rowsBySuffix = (map, suffix) => Object.entries(map || {}).filter(([key]) => key.endsWith(`:${suffix}`))
    .map(([key, value]) => [key.slice(0, -suffix.length - 1), value]).sort((a, b) => b[1] - a[1]);
  const statRows = (items) => items.length ? items.map(([name, stats]) => `<div class="control-center-row"><span>${escapeHtml(name)}</span><strong>${numeric(stats.documents)} doc · ${numeric(stats.operations)} op</strong></div>`).join('') : '<p class="muted">Nessun dato.</p>';

  function ensureCard() {
    let card = document.getElementById('firestore-operation-diagnostics-card');
    if (card) return card;
    const root = document.getElementById('control-center-content') || document.getElementById('control-center-page');
    if (!root) return null;
    card = document.createElement('section');
    card.id = 'firestore-operation-diagnostics-card';
    card.className = 'card';
    card.innerHTML = '<div class="section-head"><div><h2>🔎 Diagnostica operazioni Firestore V3</h2><p class="muted">Attribuisce letture, consegne dei listener, funzioni e schermate senza generare operazioni Firestore.</p></div><button class="btn" data-refresh>AGGIORNA</button></div><div data-summary></div><div class="actions-row"><button class="btn" data-export>SCARICA REPORT</button><button class="btn" data-reset>AZZERA OGGI</button></div>';
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
      anchor.download = `diagnostica-firestore-${value.date}-v3.json`;
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
    const topFunctions = sortedStats(value.functions).slice(0, 20);
    const topQueries = sortedStats(value.queries).slice(0, 20);
    const topScreens = sortedStats(value.screens).slice(0, 20);
    const listeners = Object.values(value.listenerInstances || {});
    const active = listeners.filter((item) => item.active).sort((a, b) => String(a.openedAt).localeCompare(String(b.openedAt)));
    const costly = listeners.slice().sort((a, b) => numeric(b.documents) - numeric(a.documents)).slice(0, 20);
    const readsByArea = rowsBySuffix(value.areas, 'read');

    card.querySelector('[data-summary]').innerHTML = `
      <div class="control-center-grid">
        <div class="control-center-row"><span>Chiamate di lettura</span><strong>${value.readOperations}</strong></div>
        <div class="control-center-row"><span>Documenti letti totali</span><strong>${value.readDocuments}</strong></div>
        <div class="control-center-row"><span>Snapshot listener consegnati</span><strong>${value.listenerDeliveries}</strong></div>
        <div class="control-center-row"><span>Documenti da listener</span><strong>${value.listenerDocuments}</strong></div>
        <div class="control-center-row"><span>Letture non attribuite</span><strong>${value.unattributedReads}</strong></div>
        <div class="control-center-row"><span>Latenza media get()</span><strong>${value.readOperations ? Math.round(value.readLatencyMsTotal / value.readOperations) : 0} ms</strong></div>
        <div class="control-center-row"><span>Latenza massima get()</span><strong>${value.readLatencyMsMax} ms</strong></div>
        <div class="control-center-row"><span>Listener attivi adesso</span><strong>${activeListeners}</strong></div>
        <div class="control-center-row"><span>Picco listener attivi</span><strong>${Math.max(value.peakActiveListeners || 0, activeListeners)}</strong></div>
        <div class="control-center-row"><span>Dettagli scartati</span><strong>${value.detailsDropped}</strong></div>
      </div>
      <details open><summary><strong>Top funzioni per documenti letti</strong></summary>${statRows(topFunctions)}</details>
      <details><summary><strong>Top query/percorso per documenti letti</strong></summary>${statRows(topQueries)}</details>
      <details><summary><strong>Top schermate per documenti letti</strong></summary>${statRows(topScreens)}</details>
      <details><summary><strong>Listener attivi adesso</strong></summary>${active.length ? active.map((item) => `<div class="control-center-row"><span>${escapeHtml(`${item.functionName} · ${item.path}`)}</span><strong>${Math.round((Date.now() - new Date(item.openedAt).getTime()) / 1000)} s</strong></div>`).join('') : '<p class="muted">Nessun listener attivo.</p>'}</details>
      <details><summary><strong>Listener più costosi</strong></summary>${costly.length ? costly.map((item) => `<div class="control-center-row"><span>${escapeHtml(`${item.functionName} · ${item.path}`)}</span><strong>${numeric(item.documents)} doc · ${numeric(item.deliveries)} snapshot</strong></div>`).join('') : '<p class="muted">Nessun listener osservato.</p>'}</details>
      <details><summary><strong>Letture one-shot per area</strong></summary>${readsByArea.length ? readsByArea.map(([name, count]) => `<div class="control-center-row"><span>${escapeHtml(name)}</span><strong>${count}</strong></div>`).join('') : '<p class="muted">Nessuna lettura osservata.</p>'}</details>
      <p class="muted">Monitor V3 attivo dal ${new Date(value.startedAt).toLocaleString('it-IT')}. Il report resta locale e non apre query o listener aggiuntivi.</p>`;
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
    activeListeners: () => activeListeners,
    liveListeners: () => Array.from(liveListeners.values()),
    version: SCRIPT_VERSION
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
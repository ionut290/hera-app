(() => {
  'use strict';

  if (window.__vargaFirestoreDiagnosticsV4) return;
  window.__vargaFirestoreDiagnosticsV4 = true;

  const SCRIPT_VERSION = '4.0.0';
  const SCHEMA_VERSION = 4;
  const PREFIX = 'varga_fs_diag_v4_';
  const EVENT_LIMIT = 1500;
  const ERROR_LIMIT = 150;
  const LIFECYCLE_LIMIT = 250;
  const DEFAULT_DAILY_THRESHOLD = 50000;
  const listenerRuntime = new Map();
  const batchRuntime = new WeakMap();
  let listenerSequence = 0;
  let patched = false;
  let renderTimer = 0;
  let lifecycleInstalled = false;

  const text = (value) => String(value ?? '').trim();
  const number = (value) => Number.isFinite(Number(value)) ? Math.max(0, Number(value)) : 0;
  const nowIso = () => new Date().toISOString();
  const nowMs = () => Date.now();
  const today = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Europe/Rome' });
  const storageKey = () => `${PREFIX}${today()}`;
  const sessionId = () => {
    try {
      return window.VargaFirestoreDiagnostics?.read?.()?.sessionId || `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
    } catch (_) {
      return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
    }
  };

  function baseTotals() {
    return {
      oneShotOperations: 0,
      oneShotPayloadDocuments: 0,
      oneShotCacheDocuments: 0,
      oneShotServerDocuments: 0,
      oneShotEstimatedReadsMin: 0,
      oneShotEstimatedReadsMax: 0,
      listenerRegistrations: 0,
      listenerUnsubscribes: 0,
      listenerDeliveries: 0,
      listenerPayloadDocuments: 0,
      listenerChangedDocuments: 0,
      listenerCacheDocuments: 0,
      listenerServerDocuments: 0,
      listenerEstimatedReadsMin: 0,
      listenerEstimatedReadsMax: 0,
      metadataOnlyDeliveries: 0,
      pendingWriteDeliveries: 0,
      emptyServerQueries: 0,
      duplicateListenerRegistrations: 0,
      quickListenerReopens: 0,
      shortLivedListeners: 0,
      writesSucceeded: 0,
      writesFailed: 0,
      deletesSucceeded: 0,
      deletesFailed: 0,
      batchesCommitted: 0,
      batchOperations: 0,
      transactionsStarted: 0,
      transactionAttempts: 0,
      transactionWriteOperationsScheduled: 0,
      transactionDeleteOperationsScheduled: 0,
      transactionsSucceeded: 0,
      transactionsFailed: 0,
      operationFailures: 0,
      offlineOperations: 0,
      cacheKnownOperations: 0,
      sourceUnknownOperations: 0
    };
  }

  function blank() {
    return {
      schemaVersion: SCHEMA_VERSION,
      scriptVersion: SCRIPT_VERSION,
      date: today(),
      sessionId: sessionId(),
      startedAt: nowIso(),
      updatedAt: nowIso(),
      initialUrl: location.href,
      totals: baseTotals(),
      methods: {},
      collections: {},
      paths: {},
      screens: {},
      callers: {},
      minuteBuckets: {},
      listeners: {},
      lastClosedByFingerprint: {},
      lifecycle: [],
      errors: [],
      events: [],
      eventsDropped: 0,
      errorsDropped: 0,
      lifecycleDropped: 0
    };
  }

  function migrate(value) {
    if (!value || value.date !== today()) return blank();
    const base = blank();
    return {
      ...base,
      ...value,
      schemaVersion: SCHEMA_VERSION,
      scriptVersion: SCRIPT_VERSION,
      totals: { ...base.totals, ...(value.totals || {}) },
      methods: value.methods && typeof value.methods === 'object' ? value.methods : {},
      collections: value.collections && typeof value.collections === 'object' ? value.collections : {},
      paths: value.paths && typeof value.paths === 'object' ? value.paths : {},
      screens: value.screens && typeof value.screens === 'object' ? value.screens : {},
      callers: value.callers && typeof value.callers === 'object' ? value.callers : {},
      minuteBuckets: value.minuteBuckets && typeof value.minuteBuckets === 'object' ? value.minuteBuckets : {},
      listeners: value.listeners && typeof value.listeners === 'object' ? value.listeners : {},
      lastClosedByFingerprint: value.lastClosedByFingerprint && typeof value.lastClosedByFingerprint === 'object' ? value.lastClosedByFingerprint : {},
      lifecycle: Array.isArray(value.lifecycle) ? value.lifecycle.slice(0, LIFECYCLE_LIMIT) : [],
      errors: Array.isArray(value.errors) ? value.errors.slice(0, ERROR_LIMIT) : [],
      events: Array.isArray(value.events) ? value.events.slice(0, EVENT_LIMIT) : []
    };
  }

  function load() {
    try {
      return migrate(JSON.parse(localStorage.getItem(storageKey()) || 'null'));
    } catch (_) {
      return blank();
    }
  }

  function trimArray(value, key, limit, droppedKey) {
    const list = Array.isArray(value[key]) ? value[key] : [];
    if (list.length > limit) {
      value[droppedKey] = number(value[droppedKey]) + list.length - limit;
      list.length = limit;
    }
    value[key] = list;
  }

  function persist(value) {
    value.updatedAt = nowIso();
    trimArray(value, 'events', EVENT_LIMIT, 'eventsDropped');
    trimArray(value, 'errors', ERROR_LIMIT, 'errorsDropped');
    trimArray(value, 'lifecycle', LIFECYCLE_LIMIT, 'lifecycleDropped');
    let events = value.events;
    while (true) {
      try {
        value.events = events;
        localStorage.setItem(storageKey(), JSON.stringify(value));
        break;
      } catch (_) {
        if (!events.length) break;
        const remove = Math.max(1, Math.ceil(events.length / 4));
        events = events.slice(0, Math.max(0, events.length - remove));
        value.eventsDropped = number(value.eventsDropped) + remove;
      }
    }
    clearTimeout(renderTimer);
    renderTimer = setTimeout(renderDashboard, 160);
  }

  function increment(object, key, amount = 1) {
    object[key] = number(object[key]) + number(amount);
  }

  function addMetric(map, key, metrics = {}) {
    const current = map[key] && typeof map[key] === 'object' ? map[key] : {};
    Object.entries(metrics).forEach(([name, amount]) => increment(current, name, amount));
    current.lastAt = nowIso();
    map[key] = current;
  }

  function callerFromStack(stack) {
    const lines = text(stack).split('\n').map((line) => line.trim());
    return lines.find((line) => line &&
      !/firestore-diagnostics-dashboard-v4\.js/i.test(line) &&
      !/firestore-operation-diagnostics\.js/i.test(line) &&
      !/firestore-safe-optimizer\.js/i.test(line) &&
      !/firebase(?:-firestore)?(?:\.js)?/i.test(line) &&
      !/new Error/i.test(line)) || 'chiamante non identificato';
  }

  function functionNameFromCaller(caller) {
    const match = text(caller).match(/(?:at\s+)?([^\s@()]+)\s*(?:@|\()/);
    const value = text(match?.[1]);
    if (!value || /https?:|anonymous|<anonymous>|chiamante/i.test(value)) return 'non attribuita';
    return value.replace(/^Object\./, '').replace(/^window\./, '');
  }

  function currentScreen() {
    try {
      const visible = Array.from(document.querySelectorAll('[data-page], .page, section[id], main[id]'))
        .find((element) => element.offsetParent !== null && getComputedStyle(element).display !== 'none');
      return visible?.dataset?.page || visible?.id || location.hash || location.pathname || 'sconosciuta';
    } catch (_) {
      return location.hash || location.pathname || 'sconosciuta';
    }
  }

  function queryDescription(query) {
    const internal = query?._query || query?.Ae || query?.je || query?._delegate?._query || null;
    const path = internal?.path?.canonicalString?.() || internal?.path?.toString?.() ||
      internal?.path?.segments?.join?.('/') || query?.path || 'query-senza-percorso';
    let canonical = '';
    try { canonical = internal?.canonicalId?.() || internal?.toString?.() || ''; } catch (_) {}
    return canonical && canonical !== '[object Object]' ? `${path}|query=${canonical}` : path;
  }

  function firstCollection(path) {
    const plain = text(path).split('|query=')[0];
    return plain.split('/').filter(Boolean)[0] || 'sconosciuta';
  }

  function snapshotInfo(snapshot, kind, previousFingerprint = '') {
    const isQuery = kind === 'query';
    const fromCache = typeof snapshot?.metadata?.fromCache === 'boolean' ? snapshot.metadata.fromCache : null;
    const hasPendingWrites = typeof snapshot?.metadata?.hasPendingWrites === 'boolean' ? snapshot.metadata.hasPendingWrites : null;
    const payloadDocuments = isQuery
      ? number(snapshot?.size)
      : (typeof snapshot?.exists === 'boolean' ? (snapshot.exists ? 1 : 0) : 0);
    let changedDocuments = 0;
    let fingerprint = previousFingerprint;

    if (isQuery) {
      try { changedDocuments = number(snapshot?.docChanges?.().length); } catch (_) { changedDocuments = payloadDocuments; }
    } else {
      try {
        const updateTime = snapshot?.updateTime?.toMillis?.() || snapshot?.updateTime?.seconds || '';
        const data = snapshot?.exists ? JSON.stringify(snapshot.data?.() ?? null) : '__missing__';
        fingerprint = `${snapshot?.id || ''}|${updateTime}|${data}`;
        changedDocuments = previousFingerprint && previousFingerprint === fingerprint ? 0 : 1;
      } catch (_) {
        changedDocuments = 1;
      }
    }

    return {
      fromCache,
      hasPendingWrites,
      payloadDocuments,
      changedDocuments,
      fingerprint
    };
  }

  function serverEstimate(info, kind, initialDelivery) {
    if (info.fromCache === true) return { min: 0, max: 0, emptyQuery: false };
    if (info.hasPendingWrites === true && !initialDelivery) {
      const max = kind === 'query' ? Math.max(0, info.payloadDocuments) : 1;
      return { min: 0, max, emptyQuery: false };
    }
    if (info.fromCache === null) {
      const max = kind === 'query' ? Math.max(1, info.payloadDocuments) : 1;
      return { min: 0, max, emptyQuery: kind === 'query' && info.payloadDocuments === 0 };
    }
    if (kind === 'doc') {
      const min = initialDelivery || info.changedDocuments > 0 ? 1 : 0;
      return { min, max: 1, emptyQuery: false };
    }
    const min = initialDelivery ? Math.max(1, info.payloadDocuments) : info.changedDocuments;
    const max = Math.max(min, info.payloadDocuments || (initialDelivery ? 1 : 0));
    return { min, max, emptyQuery: initialDelivery && info.payloadDocuments === 0 };
  }

  function minuteKey(at = new Date()) {
    return at.toISOString().slice(0, 16);
  }

  function recordOperation(payload) {
    const value = load();
    const at = payload.at || nowIso();
    const path = text(payload.path) || 'sconosciuto';
    const collection = firstCollection(path);
    const screen = payload.screen || currentScreen();
    const caller = payload.caller || 'chiamante non identificato';
    const method = payload.method || 'sconosciuto';
    const metrics = {
      operations: 1,
      payloadDocuments: number(payload.payloadDocuments),
      changedDocuments: number(payload.changedDocuments),
      cacheDocuments: number(payload.cacheDocuments),
      serverDocuments: number(payload.serverDocuments),
      estimatedReadsMin: number(payload.estimatedReadsMin),
      estimatedReadsMax: number(payload.estimatedReadsMax),
      durationMsTotal: number(payload.durationMs),
      failures: payload.failed ? 1 : 0
    };

    addMetric(value.methods, method, metrics);
    addMetric(value.collections, collection, metrics);
    addMetric(value.paths, path, metrics);
    addMetric(value.screens, screen, metrics);
    addMetric(value.callers, caller, metrics);
    addMetric(value.minuteBuckets, minuteKey(new Date(at)), metrics);

    value.events.unshift({
      at,
      type: payload.type,
      method,
      path,
      collection,
      screen,
      caller,
      functionName: functionNameFromCaller(caller),
      durationMs: number(payload.durationMs),
      payloadDocuments: number(payload.payloadDocuments),
      changedDocuments: number(payload.changedDocuments),
      fromCache: payload.fromCache ?? null,
      hasPendingWrites: payload.hasPendingWrites ?? null,
      estimatedReadsMin: number(payload.estimatedReadsMin),
      estimatedReadsMax: number(payload.estimatedReadsMax),
      listenerId: payload.listenerId || null,
      deliveryNumber: payload.deliveryNumber || null,
      source: payload.source || 'default',
      online: navigator.onLine,
      failed: Boolean(payload.failed),
      errorCode: payload.errorCode || null
    });
    persist(value);
  }

  function recordFailure(method, path, error, context = {}) {
    const value = load();
    const entry = {
      at: nowIso(),
      method,
      path: text(path) || 'sconosciuto',
      screen: context.screen || currentScreen(),
      caller: context.caller || 'chiamante non identificato',
      code: text(error?.code || ''),
      name: text(error?.name || ''),
      message: text(error?.message || error || 'Errore sconosciuto').slice(0, 500),
      online: navigator.onLine
    };
    value.errors.unshift(entry);
    increment(value.totals, 'operationFailures');
    if (!navigator.onLine) increment(value.totals, 'offlineOperations');
    persist(value);
    recordOperation({
      type: context.type || 'error', method, path, screen: entry.screen, caller: entry.caller,
      durationMs: context.durationMs, failed: true, errorCode: entry.code
    });
  }

  function promiseWithTelemetry(result, handlers) {
    if (!result?.then) return result;
    return result.then((value) => {
      handlers.success?.(value);
      return value;
    }, (error) => {
      handlers.failure?.(error);
      throw error;
    });
  }

  function wrap(prototype, name, factory) {
    if (!prototype || typeof prototype[name] !== 'function' || prototype[name].__vargaDiagV4) return;
    const original = prototype[name];
    const wrapped = factory(original);
    Object.defineProperty(wrapped, '__vargaDiagV4', { value: true });
    Object.defineProperty(wrapped, '__vargaOriginalV4', { value: original });
    prototype[name] = wrapped;
  }

  function wrapSnapshotCallbacks(args, onNext, onError) {
    const wrapped = args.slice();
    const observerIndex = wrapped.findIndex((arg) => arg && typeof arg === 'object' &&
      (typeof arg.next === 'function' || typeof arg.error === 'function'));
    if (observerIndex >= 0) {
      const observer = wrapped[observerIndex];
      wrapped[observerIndex] = {
        ...observer,
        next(snapshot) {
          onNext(snapshot);
          return observer.next?.call(observer, snapshot);
        },
        error(error) {
          onError(error);
          return observer.error?.call(observer, error);
        }
      };
      return wrapped;
    }

    const functionIndexes = wrapped.map((arg, index) => typeof arg === 'function' ? index : -1).filter((index) => index >= 0);
    const nextIndex = functionIndexes[0];
    const errorIndex = functionIndexes[1];
    if (nextIndex !== undefined) {
      const next = wrapped[nextIndex];
      wrapped[nextIndex] = function enhancedNext(snapshot) {
        onNext(snapshot);
        return next.apply(this, arguments);
      };
    }
    if (errorIndex !== undefined) {
      const errorCallback = wrapped[errorIndex];
      wrapped[errorIndex] = function enhancedError(error) {
        onError(error);
        return errorCallback.apply(this, arguments);
      };
    }
    return wrapped;
  }

  function listenerWrapper(original, method, kind, pathFactory) {
    return function (...args) {
      const stack = new Error().stack;
      const caller = callerFromStack(stack);
      const screen = currentScreen();
      const path = pathFactory(this);
      const listenerId = `V4-${Date.now().toString(36)}-${++listenerSequence}`;
      const fingerprint = `${method}|${path}`;
      const openedAtMs = nowMs();
      let deliveryNumber = 0;
      let lastSnapshotFingerprint = '';

      const state = load();
      const activeDuplicates = Object.values(state.listeners || {}).filter((item) => item.active && item.fingerprint === fingerprint).length;
      increment(state.totals, 'listenerRegistrations');
      if (activeDuplicates > 0) increment(state.totals, 'duplicateListenerRegistrations');
      const previousClose = number(state.lastClosedByFingerprint?.[fingerprint]);
      if (previousClose && openedAtMs - previousClose < 2000) increment(state.totals, 'quickListenerReopens');
      state.listeners[listenerId] = {
        id: listenerId,
        method,
        kind,
        path,
        collection: firstCollection(path),
        caller,
        functionName: functionNameFromCaller(caller),
        screen,
        fingerprint,
        openedAt: new Date(openedAtMs).toISOString(),
        closedAt: null,
        durationMs: null,
        active: true,
        deliveries: 0,
        payloadDocuments: 0,
        changedDocuments: 0,
        cacheDocuments: 0,
        serverDocuments: 0,
        estimatedReadsMin: 0,
        estimatedReadsMax: 0,
        duplicateAtOpen: activeDuplicates
      };
      persist(state);
      listenerRuntime.set(listenerId, { fingerprint, openedAtMs });

      const wrappedArgs = wrapSnapshotCallbacks(args, (snapshot) => {
        deliveryNumber += 1;
        const info = snapshotInfo(snapshot, kind, lastSnapshotFingerprint);
        lastSnapshotFingerprint = info.fingerprint;
        const estimate = serverEstimate(info, kind, deliveryNumber === 1);
        const value = load();
        const listener = value.listeners[listenerId] || {};
        listener.deliveries = number(listener.deliveries) + 1;
        listener.payloadDocuments = number(listener.payloadDocuments) + info.payloadDocuments;
        listener.changedDocuments = number(listener.changedDocuments) + info.changedDocuments;
        listener.cacheDocuments = number(listener.cacheDocuments) + (info.fromCache === true ? info.payloadDocuments : 0);
        listener.serverDocuments = number(listener.serverDocuments) + (info.fromCache === false ? info.payloadDocuments : 0);
        listener.estimatedReadsMin = number(listener.estimatedReadsMin) + estimate.min;
        listener.estimatedReadsMax = number(listener.estimatedReadsMax) + estimate.max;
        listener.lastDeliveryAt = nowIso();
        listener.lastFromCache = info.fromCache;
        value.listeners[listenerId] = listener;
        increment(value.totals, 'listenerDeliveries');
        increment(value.totals, 'listenerPayloadDocuments', info.payloadDocuments);
        increment(value.totals, 'listenerChangedDocuments', info.changedDocuments);
        increment(value.totals, 'listenerEstimatedReadsMin', estimate.min);
        increment(value.totals, 'listenerEstimatedReadsMax', estimate.max);
        if (info.fromCache === true) increment(value.totals, 'listenerCacheDocuments', info.payloadDocuments);
        if (info.fromCache === false) increment(value.totals, 'listenerServerDocuments', info.payloadDocuments);
        if (info.fromCache !== null) increment(value.totals, 'cacheKnownOperations');
        else increment(value.totals, 'sourceUnknownOperations');
        if (info.fromCache === false && estimate.emptyQuery) increment(value.totals, 'emptyServerQueries');
        if (deliveryNumber > 1 && info.changedDocuments === 0) increment(value.totals, 'metadataOnlyDeliveries');
        if (info.hasPendingWrites === true) increment(value.totals, 'pendingWriteDeliveries');
        persist(value);

        recordOperation({
          type: 'listener-delivery', method: `${method}.delivery`, path, screen, caller,
          durationMs: deliveryNumber === 1 ? nowMs() - openedAtMs : 0,
          payloadDocuments: info.payloadDocuments,
          changedDocuments: info.changedDocuments,
          cacheDocuments: info.fromCache === true ? info.payloadDocuments : 0,
          serverDocuments: info.fromCache === false ? info.payloadDocuments : 0,
          fromCache: info.fromCache,
          hasPendingWrites: info.hasPendingWrites,
          estimatedReadsMin: estimate.min,
          estimatedReadsMax: estimate.max,
          listenerId,
          deliveryNumber
        });
      }, (error) => {
        recordFailure(`${method}.listener`, path, error, { type: 'listener-error', screen, caller, durationMs: nowMs() - openedAtMs });
      });

      let unsubscribe;
      try {
        unsubscribe = original.apply(this, wrappedArgs);
      } catch (error) {
        recordFailure(`${method}.open`, path, error, { type: 'listener-open-error', screen, caller });
        throw error;
      }
      if (typeof unsubscribe !== 'function') return unsubscribe;

      let closed = false;
      return function diagnosticV4Unsubscribe() {
        if (!closed) {
          closed = true;
          const closedAtMs = nowMs();
          const value = load();
          const listener = value.listeners[listenerId] || {};
          listener.active = false;
          listener.closedAt = new Date(closedAtMs).toISOString();
          listener.durationMs = closedAtMs - openedAtMs;
          value.listeners[listenerId] = listener;
          value.lastClosedByFingerprint[fingerprint] = closedAtMs;
          increment(value.totals, 'listenerUnsubscribes');
          if (closedAtMs - openedAtMs < 2000) increment(value.totals, 'shortLivedListeners');
          persist(value);
          listenerRuntime.delete(listenerId);
        }
        return unsubscribe.apply(this, arguments);
      };
    };
  }

  function oneShotWrapper(original, method, kind, pathFactory) {
    return function (...args) {
      const started = nowMs();
      const stack = new Error().stack;
      const caller = callerFromStack(stack);
      const screen = currentScreen();
      const path = pathFactory(this, args);
      const source = text(args?.[0]?.source || 'default');
      let result;
      try {
        result = original.apply(this, args);
      } catch (error) {
        recordFailure(method, path, error, { type: 'read-error', screen, caller, durationMs: nowMs() - started });
        throw error;
      }
      return promiseWithTelemetry(result, {
        success: (snapshot) => {
          const info = snapshotInfo(snapshot, kind);
          const estimate = serverEstimate(info, kind, true);
          const value = load();
          increment(value.totals, 'oneShotOperations');
          increment(value.totals, 'oneShotPayloadDocuments', info.payloadDocuments);
          increment(value.totals, 'oneShotEstimatedReadsMin', estimate.min);
          increment(value.totals, 'oneShotEstimatedReadsMax', estimate.max);
          if (info.fromCache === true) increment(value.totals, 'oneShotCacheDocuments', info.payloadDocuments);
          if (info.fromCache === false) increment(value.totals, 'oneShotServerDocuments', info.payloadDocuments);
          if (info.fromCache !== null) increment(value.totals, 'cacheKnownOperations');
          else increment(value.totals, 'sourceUnknownOperations');
          if (info.fromCache === false && estimate.emptyQuery) increment(value.totals, 'emptyServerQueries');
          if (!navigator.onLine) increment(value.totals, 'offlineOperations');
          persist(value);
          recordOperation({
            type: 'read', method, path, screen, caller, source,
            durationMs: nowMs() - started,
            payloadDocuments: info.payloadDocuments,
            changedDocuments: info.changedDocuments,
            cacheDocuments: info.fromCache === true ? info.payloadDocuments : 0,
            serverDocuments: info.fromCache === false ? info.payloadDocuments : 0,
            fromCache: info.fromCache,
            hasPendingWrites: info.hasPendingWrites,
            estimatedReadsMin: estimate.min,
            estimatedReadsMax: estimate.max
          });
        },
        failure: (error) => recordFailure(method, path, error, { type: 'read-error', screen, caller, durationMs: nowMs() - started })
      });
    };
  }

  function writeWrapper(original, method, pathFactory, deleteOperation = false) {
    return function (...args) {
      const started = nowMs();
      const stack = new Error().stack;
      const caller = callerFromStack(stack);
      const screen = currentScreen();
      const path = pathFactory(this, args);
      let result;
      try {
        result = original.apply(this, args);
      } catch (error) {
        const value = load();
        increment(value.totals, deleteOperation ? 'deletesFailed' : 'writesFailed');
        persist(value);
        recordFailure(method, path, error, { type: 'write-error', screen, caller, durationMs: nowMs() - started });
        throw error;
      }
      return promiseWithTelemetry(result, {
        success: () => {
          const value = load();
          increment(value.totals, deleteOperation ? 'deletesSucceeded' : 'writesSucceeded');
          persist(value);
          recordOperation({ type: deleteOperation ? 'delete' : 'write', method, path, screen, caller, durationMs: nowMs() - started });
        },
        failure: (error) => {
          const value = load();
          increment(value.totals, deleteOperation ? 'deletesFailed' : 'writesFailed');
          persist(value);
          recordFailure(method, path, error, { type: 'write-error', screen, caller, durationMs: nowMs() - started });
        }
      });
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
    const Firestore = firestore.Firestore?.prototype;

    wrap(DocumentReference, 'get', (original) => oneShotWrapper(original, 'doc.get', 'doc', (reference) => reference.path));
    wrap(Query, 'get', (original) => oneShotWrapper(original, 'query.get', 'query', queryDescription));
    wrap(DocumentReference, 'onSnapshot', (original) => listenerWrapper(original, 'doc.onSnapshot', 'doc', (reference) => reference.path));
    wrap(Query, 'onSnapshot', (original) => listenerWrapper(original, 'query.onSnapshot', 'query', queryDescription));

    wrap(DocumentReference, 'set', (original) => writeWrapper(original, 'doc.set', (reference) => reference.path));
    wrap(DocumentReference, 'update', (original) => writeWrapper(original, 'doc.update', (reference) => reference.path));
    wrap(DocumentReference, 'delete', (original) => writeWrapper(original, 'doc.delete', (reference) => reference.path, true));
    wrap(CollectionReference, 'add', (original) => writeWrapper(original, 'collection.add', (reference) => reference.path));

    ['set', 'update', 'delete'].forEach((method) => wrap(WriteBatch, method, (original) => function (reference, ...args) {
      const operations = batchRuntime.get(this) || [];
      operations.push({ method, path: reference?.path || 'batch' });
      batchRuntime.set(this, operations);
      return original.call(this, reference, ...args);
    }));
    wrap(WriteBatch, 'commit', (original) => function (...args) {
      const started = nowMs();
      const operations = (batchRuntime.get(this) || []).slice();
      const result = original.apply(this, args);
      return promiseWithTelemetry(result, {
        success: () => {
          const value = load();
          increment(value.totals, 'batchesCommitted');
          increment(value.totals, 'batchOperations', operations.length);
          operations.forEach((item) => increment(value.totals, item.method === 'delete' ? 'deletesSucceeded' : 'writesSucceeded'));
          persist(value);
          recordOperation({ type: 'batch', method: 'batch.commit', path: operations.map((item) => item.path).join(',') || 'batch', durationMs: nowMs() - started });
          batchRuntime.delete(this);
        },
        failure: (error) => recordFailure('batch.commit', 'batch', error, { type: 'batch-error', durationMs: nowMs() - started })
      });
    });

    wrap(Transaction, 'get', (original) => oneShotWrapper(original, 'transaction.get', 'doc', (_transaction, args) => args?.[0]?.path || 'transaction'));
    ['set', 'update'].forEach((method) => wrap(Transaction, method, (original) => function (reference, ...args) {
      const result = original.call(this, reference, ...args);
      const value = load();
      increment(value.totals, 'transactionWriteOperationsScheduled');
      persist(value);
      recordOperation({ type: 'transaction-write-scheduled', method: `transaction.${method}`, path: reference?.path || 'transaction' });
      return result;
    }));
    wrap(Transaction, 'delete', (original) => function (reference, ...args) {
      const result = original.call(this, reference, ...args);
      const value = load();
      increment(value.totals, 'transactionDeleteOperationsScheduled');
      persist(value);
      recordOperation({ type: 'transaction-delete-scheduled', method: 'transaction.delete', path: reference?.path || 'transaction' });
      return result;
    });

    ['enableNetwork', 'disableNetwork', 'clearPersistence', 'terminate'].forEach((method) => wrap(Firestore, method, (original) => function (...args) {
      const started = nowMs();
      const result = original.apply(this, args);
      return promiseWithTelemetry(result, {
        success: () => recordOperation({ type: 'firestore-control', method: `firestore.${method}`, path: 'firestore-runtime', durationMs: nowMs() - started }),
        failure: (error) => recordFailure(`firestore.${method}`, 'firestore-runtime', error, { type: 'firestore-control-error', durationMs: nowMs() - started })
      });
    }));

    wrap(Firestore, 'runTransaction', (original) => function (...args) {
      const started = nowMs();
      const value = load();
      increment(value.totals, 'transactionsStarted');
      persist(value);
      const wrappedArgs = args.slice();
      const updateIndex = wrappedArgs.findIndex((arg) => typeof arg === 'function');
      if (updateIndex >= 0) {
        const updateFunction = wrappedArgs[updateIndex];
        wrappedArgs[updateIndex] = function diagnosticTransactionAttempt(...callbackArgs) {
          const attemptState = load();
          increment(attemptState.totals, 'transactionAttempts');
          persist(attemptState);
          return updateFunction.apply(this, callbackArgs);
        };
      }
      const result = original.apply(this, wrappedArgs);
      return promiseWithTelemetry(result, {
        success: () => {
          const state = load();
          increment(state.totals, 'transactionsSucceeded');
          persist(state);
          recordOperation({ type: 'transaction', method: 'firestore.runTransaction', path: 'transaction', durationMs: nowMs() - started });
        },
        failure: (error) => {
          const state = load();
          increment(state.totals, 'transactionsFailed');
          persist(state);
          recordFailure('firestore.runTransaction', 'transaction', error, { type: 'transaction-error', durationMs: nowMs() - started });
        }
      });
    });

    patched = true;
    return true;
  }

  function contextSnapshot() {
    const connection = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
    const memory = performance?.memory;
    return {
      capturedAt: nowIso(),
      url: location.href,
      pathname: location.pathname,
      hash: location.hash,
      referrer: document.referrer,
      title: document.title,
      language: navigator.language,
      languages: navigator.languages,
      userAgent: navigator.userAgent,
      platform: navigator.platform,
      online: navigator.onLine,
      visibilityState: document.visibilityState,
      viewport: { width: innerWidth, height: innerHeight, devicePixelRatio: devicePixelRatio || 1 },
      screen: { width: window.screen?.width, height: window.screen?.height, colorDepth: window.screen?.colorDepth },
      connection: connection ? {
        effectiveType: connection.effectiveType,
        downlinkMbps: connection.downlink,
        rttMs: connection.rtt,
        saveData: connection.saveData,
        type: connection.type
      } : null,
      memory: memory ? {
        usedJsHeapSize: memory.usedJSHeapSize,
        totalJsHeapSize: memory.totalJSHeapSize,
        jsHeapSizeLimit: memory.jsHeapSizeLimit
      } : null,
      serviceWorkerControlled: Boolean(navigator.serviceWorker?.controller),
      firebaseProjectId: window.firebaseConfig?.projectId || '',
      buildVersion: window.APP_VERSION || window.BUILD_VERSION || ''
    };
  }

  function activeDuplicateGroups(listeners) {
    const groups = {};
    Object.values(listeners || {}).filter((item) => item.active).forEach((item) => {
      const key = item.fingerprint || `${item.method}|${item.path}`;
      (groups[key] ||= []).push(item);
    });
    return Object.entries(groups).filter(([, items]) => items.length > 1).map(([fingerprint, items]) => ({
      fingerprint,
      count: items.length,
      path: items[0]?.path,
      method: items[0]?.method,
      screens: [...new Set(items.map((item) => item.screen))],
      callers: [...new Set(items.map((item) => item.functionName || item.caller))]
    })).sort((a, b) => b.count - a.count);
  }

  function recommendations(value, legacy) {
    const totals = value.totals || {};
    const estimated = number(totals.oneShotEstimatedReadsMin) + number(totals.listenerEstimatedReadsMin);
    const payload = number(totals.oneShotPayloadDocuments) + number(totals.listenerPayloadDocuments);
    const duplicates = activeDuplicateGroups(value.listeners);
    const active = Object.values(value.listeners || {}).filter((item) => item.active);
    const collections = Object.entries(value.collections || {}).sort((a, b) => number(b[1]?.estimatedReadsMin) - number(a[1]?.estimatedReadsMin));
    const issues = [];

    if (duplicates.length) issues.push({ severity: 'high', code: 'duplicate-listeners', message: `${duplicates.length} gruppi di listener fisici duplicati sono attivi.` });
    if (active.length > 25) issues.push({ severity: 'high', code: 'many-active-listeners', message: `${active.length} listener fisici attivi: verificare l’inizializzazione delle schermate.` });
    else if (active.length > 15) issues.push({ severity: 'medium', code: 'active-listeners', message: `${active.length} listener attivi: numero elevato per una singola schermata.` });
    if (number(totals.quickListenerReopens) > 0) issues.push({ severity: 'medium', code: 'quick-reopen', message: `${number(totals.quickListenerReopens)} listener riaperti entro 2 secondi dalla chiusura.` });
    if (number(totals.shortLivedListeners) > 3) issues.push({ severity: 'medium', code: 'listener-churn', message: `${number(totals.shortLivedListeners)} listener durati meno di 2 secondi.` });
    if (number(totals.operationFailures) > 0) issues.push({ severity: 'high', code: 'errors', message: `${number(totals.operationFailures)} operazioni Firestore fallite.` });
    if (number(totals.sourceUnknownOperations) > 0) issues.push({ severity: 'low', code: 'unknown-source', message: `${number(totals.sourceUnknownOperations)} snapshot senza informazione cache/server.` });
    if (collections[0] && estimated > 0 && number(collections[0][1]?.estimatedReadsMin) / estimated >= 0.5) {
      issues.push({ severity: 'high', code: 'hot-collection', message: `${collections[0][0]} genera almeno il 50% delle letture server stimate.` });
    }
    if (number(legacy?.unattributedReads) > 0) issues.push({ severity: 'medium', code: 'unattributed', message: `${number(legacy.unattributedReads)} documenti nel monitor V3 non hanno un chiamante attribuito.` });
    if (number(legacy?.detailsDropped) > 0 || number(value.eventsDropped) > 0) issues.push({ severity: 'low', code: 'truncated', message: 'Una parte dei dettagli è stata rimossa per rispettare lo spazio locale.' });
    if (payload > estimated * 1.5 && payload > 20) issues.push({ severity: 'info', code: 'payload-vs-billing', message: 'Molti documenti visualizzati provengono da cache o snapshot completi: il payload non coincide con le letture fatturabili.' });
    if (!issues.length) issues.push({ severity: 'ok', code: 'healthy', message: 'Nessuna anomalia evidente nella sessione osservata.' });
    return issues;
  }

  function percentile(values, ratio) {
    const sorted = values.map(number).filter((item) => item >= 0).sort((a, b) => a - b);
    if (!sorted.length) return 0;
    return sorted[Math.min(sorted.length - 1, Math.max(0, Math.ceil(sorted.length * ratio) - 1))];
  }

  function latencySummary(events) {
    const reads = (events || []).filter((item) => item.type === 'read').map((item) => item.durationMs);
    const listenerInitial = (events || []).filter((item) => item.type === 'listener-delivery' && item.deliveryNumber === 1).map((item) => item.durationMs);
    return {
      oneShotReadMs: { count: reads.length, p50: percentile(reads, .5), p95: percentile(reads, .95), p99: percentile(reads, .99), max: Math.max(0, ...reads) },
      listenerInitialDeliveryMs: { count: listenerInitial.length, p50: percentile(listenerInitial, .5), p95: percentile(listenerInitial, .95), p99: percentile(listenerInitial, .99), max: Math.max(0, ...listenerInitial) }
    };
  }

  async function buildReport() {
    const value = load();
    const legacy = window.VargaFirestoreOptimizerDiagnostics?.read?.() || window.VargaFirestoreDiagnostics?.read?.() || {};
    let storage = null;
    try {
      const estimate = await navigator.storage?.estimate?.();
      if (estimate) storage = { usageBytes: estimate.usage, quotaBytes: estimate.quota, usagePercent: estimate.quota ? Math.round((estimate.usage / estimate.quota) * 10000) / 100 : null };
    } catch (_) {}

    const active = Object.values(value.listeners || {}).filter((item) => item.active);
    const totals = value.totals || {};
    const estimatedReadsMin = number(totals.oneShotEstimatedReadsMin) + number(totals.listenerEstimatedReadsMin);
    const estimatedReadsMax = number(totals.oneShotEstimatedReadsMax) + number(totals.listenerEstimatedReadsMax);
    const cacheDocuments = number(totals.oneShotCacheDocuments) + number(totals.listenerCacheDocuments);
    const serverDocuments = number(totals.oneShotServerDocuments) + number(totals.listenerServerDocuments);
    const payloadDocuments = number(totals.oneShotPayloadDocuments) + number(totals.listenerPayloadDocuments);
    const durationMs = Math.max(1, nowMs() - new Date(value.startedAt).getTime());
    const threshold = number(localStorage.getItem('varga_fs_diag_daily_threshold')) || DEFAULT_DAILY_THRESHOLD;

    return {
      ...legacy,
      reportSchemaVersion: SCHEMA_VERSION,
      reportGeneratedAt: nowIso(),
      diagnosticsV4: {
        schemaVersion: SCHEMA_VERSION,
        scriptVersion: SCRIPT_VERSION,
        date: value.date,
        sessionId: value.sessionId,
        startedAt: value.startedAt,
        updatedAt: value.updatedAt,
        sessionDurationMs: durationMs,
        context: { ...contextSnapshot(), storage },
        totals,
        estimates: {
          estimatedServerReadsMin: estimatedReadsMin,
          estimatedServerReadsMax: estimatedReadsMax,
          cacheDocuments,
          serverPayloadDocuments: serverDocuments,
          totalPayloadDocuments: payloadDocuments,
          readsPerMinuteMin: Math.round((estimatedReadsMin / durationMs) * 600000) / 10,
          projectedReadsPer24hAtCurrentRateMin: Math.round((estimatedReadsMin / durationMs) * 86400000),
          configuredDailyThreshold: threshold,
          thresholdProgressPercent: Math.round((estimatedReadsMin / threshold) * 10000) / 100,
          note: 'Stima tecnica lato client. Il conteggio ufficiale resta quello della console Google Cloud/Firestore.'
        },
        latency: latencySummary(value.events),
        quality: {
          cacheKnownOperations: number(totals.cacheKnownOperations),
          sourceUnknownOperations: number(totals.sourceUnknownOperations),
          eventsDropped: number(value.eventsDropped),
          errorsDropped: number(value.errorsDropped),
          lifecycleDropped: number(value.lifecycleDropped),
          listenerDeliveryMinUsesDocumentChanges: true,
          emptyServerQueriesCountedAsMinimumOneRead: true
        },
        activeListenerCount: active.length,
        duplicateActiveListenerGroups: activeDuplicateGroups(value.listeners),
        recommendations: recommendations(value, legacy),
        collections: value.collections,
        paths: value.paths,
        methods: value.methods,
        screens: value.screens,
        callers: value.callers,
        minuteBuckets: value.minuteBuckets,
        listeners: value.listeners,
        lifecycle: value.lifecycle,
        errors: value.errors,
        events: value.events
      }
    };
  }

  async function downloadReport() {
    const report = await buildReport();
    const date = report?.diagnosticsV4?.date || today();
    const blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `diagnostica-firestore-completa-${date}-v4.json`;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function installStyles() {
    if (document.getElementById('varga-firestore-v4-styles')) return;
    const style = document.createElement('style');
    style.id = 'varga-firestore-v4-styles';
    style.textContent = `
      .fs-v4-dashboard{margin:14px 0 18px;padding:14px;border:1px solid rgba(100,116,139,.25);border-radius:18px;background:linear-gradient(145deg,rgba(15,23,42,.04),rgba(148,163,184,.08));}
      .fs-v4-head{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:12px}.fs-v4-head h3{margin:0;font-size:1.05rem}.fs-v4-actions{display:flex;gap:8px;flex-wrap:wrap}
      .fs-v4-status{display:grid;grid-template-columns:auto 1fr;gap:12px;align-items:center;padding:12px;border-radius:14px;background:var(--card-bg,#fff);box-shadow:0 4px 18px rgba(15,23,42,.07);margin-bottom:12px}
      .fs-v4-light{width:42px;height:42px;border-radius:50%;box-shadow:0 0 0 7px rgba(148,163,184,.15)}.fs-v4-light.ok{background:#16a34a}.fs-v4-light.medium{background:#eab308}.fs-v4-light.high{background:#dc2626}.fs-v4-status strong{display:block;font-size:1rem}.fs-v4-status small{display:block;opacity:.75;margin-top:3px}
      .fs-v4-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:9px;margin-bottom:12px}.fs-v4-kpi{padding:11px;border-radius:13px;background:var(--card-bg,#fff);box-shadow:0 3px 12px rgba(15,23,42,.06)}.fs-v4-kpi span{display:block;font-size:.72rem;opacity:.7;text-transform:uppercase;letter-spacing:.03em}.fs-v4-kpi strong{display:block;font-size:1.25rem;margin-top:4px}.fs-v4-kpi small{font-size:.68rem;opacity:.66}
      .fs-v4-grid{display:grid;grid-template-columns:minmax(190px,.8fr) minmax(260px,1.5fr);gap:12px}@media(max-width:720px){.fs-v4-grid{grid-template-columns:1fr}}
      .fs-v4-panel{padding:12px;border-radius:14px;background:var(--card-bg,#fff);box-shadow:0 3px 12px rgba(15,23,42,.06)}.fs-v4-panel h4{margin:0 0 10px;font-size:.9rem}
      .fs-v4-donut-wrap{display:flex;align-items:center;gap:14px}.fs-v4-donut{width:96px;height:96px;border-radius:50%;position:relative;flex:0 0 auto}.fs-v4-donut:after{content:'';position:absolute;inset:17px;border-radius:50%;background:var(--card-bg,#fff)}.fs-v4-legend{font-size:.78rem}.fs-v4-legend div{display:flex;align-items:center;gap:6px;margin:7px 0}.fs-v4-dot{width:10px;height:10px;border-radius:50%}.fs-v4-dot.server{background:#2563eb}.fs-v4-dot.cache{background:#16a34a}.fs-v4-dot.unknown{background:#94a3b8}
      .fs-v4-bar-row{display:grid;grid-template-columns:minmax(90px,1fr) minmax(110px,2fr) auto;gap:8px;align-items:center;margin:8px 0;font-size:.76rem}.fs-v4-bar-label{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.fs-v4-bar-track{height:9px;border-radius:99px;background:rgba(148,163,184,.22);overflow:hidden}.fs-v4-bar-fill{height:100%;border-radius:99px;background:linear-gradient(90deg,#2563eb,#7c3aed)}
      .fs-v4-timeline{display:flex;align-items:flex-end;gap:4px;height:72px;padding-top:6px}.fs-v4-timeline span{flex:1;min-width:5px;border-radius:4px 4px 0 0;background:linear-gradient(180deg,#7c3aed,#2563eb);position:relative}.fs-v4-timeline span:hover:after{content:attr(data-label);position:absolute;bottom:100%;left:50%;transform:translateX(-50%);white-space:nowrap;background:#0f172a;color:#fff;padding:4px 6px;border-radius:6px;font-size:.65rem;z-index:3}
      .fs-v4-alerts{margin-top:12px;display:grid;gap:7px}.fs-v4-alert{padding:9px 10px;border-radius:10px;font-size:.78rem;border-left:4px solid #94a3b8;background:rgba(148,163,184,.1)}.fs-v4-alert.high{border-color:#dc2626;background:rgba(220,38,38,.08)}.fs-v4-alert.medium{border-color:#eab308;background:rgba(234,179,8,.09)}.fs-v4-alert.ok{border-color:#16a34a;background:rgba(22,163,74,.08)}.fs-v4-alert.info{border-color:#2563eb;background:rgba(37,99,235,.08)}
      .fs-v4-progress{height:11px;border-radius:99px;overflow:hidden;background:rgba(148,163,184,.22);margin-top:7px}.fs-v4-progress span{display:block;height:100%;border-radius:99px;background:linear-gradient(90deg,#16a34a,#eab308,#dc2626)}.fs-v4-note{font-size:.7rem;opacity:.66;margin:9px 0 0}
    `;
    document.head.appendChild(style);
  }

  function formatNumber(value) {
    return new Intl.NumberFormat('it-IT', { maximumFractionDigits: 1 }).format(number(value));
  }

  function escapeHtml(value) {
    return text(value).replace(/[&<>"']/g, (character) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[character]));
  }

  function renderBars(map, metric = 'estimatedReadsMin', limit = 6) {
    const items = Object.entries(map || {}).sort((a, b) => number(b[1]?.[metric]) - number(a[1]?.[metric])).slice(0, limit);
    const max = Math.max(1, ...items.map(([, stats]) => number(stats?.[metric])));
    if (!items.length) return '<p class="muted">Dati in raccolta.</p>';
    return items.map(([name, stats]) => {
      const value = number(stats?.[metric]);
      return `<div class="fs-v4-bar-row"><span class="fs-v4-bar-label" title="${escapeHtml(name)}">${escapeHtml(name)}</span><span class="fs-v4-bar-track"><span class="fs-v4-bar-fill" style="width:${Math.max(2, Math.round((value / max) * 100))}%"></span></span><strong>${formatNumber(value)}</strong></div>`;
    }).join('');
  }

  function renderDashboard() {
    installStyles();
    const card = document.getElementById('firestore-operation-diagnostics-card');
    const summary = card?.querySelector('[data-summary]');
    if (!summary) return false;

    let root = summary.querySelector('[data-firestore-v4-dashboard]');
    if (!root) {
      root = document.createElement('section');
      root.dataset.firestoreV4Dashboard = 'true';
      root.className = 'fs-v4-dashboard';
      summary.prepend(root);
    }

    const value = load();
    const legacy = window.VargaFirestoreDiagnostics?.read?.() || {};
    const optimizer = window.VargaFirestoreOptimizerDiagnostics?.optimizerState?.() || {};
    const coalescer = window.VargaFirestoreOptimizerDiagnostics?.inflightCoalescerState?.() || {};
    const totals = value.totals || {};
    const estimatedMin = number(totals.oneShotEstimatedReadsMin) + number(totals.listenerEstimatedReadsMin);
    const estimatedMax = number(totals.oneShotEstimatedReadsMax) + number(totals.listenerEstimatedReadsMax);
    const cacheDocs = number(totals.oneShotCacheDocuments) + number(totals.listenerCacheDocuments);
    const serverDocs = number(totals.oneShotServerDocuments) + number(totals.listenerServerDocuments);
    const payloadDocs = number(totals.oneShotPayloadDocuments) + number(totals.listenerPayloadDocuments);
    const unknownDocs = Math.max(0, payloadDocs - cacheDocs - serverDocs);
    const chartTotal = Math.max(1, cacheDocs + serverDocs + unknownDocs);
    const serverPct = Math.round((serverDocs / chartTotal) * 100);
    const cachePct = Math.round((cacheDocs / chartTotal) * 100);
    const active = Object.values(value.listeners || {}).filter((item) => item.active);
    const duplicateGroups = activeDuplicateGroups(value.listeners);
    const avoided = number(optimizer?.stats?.reusedDeviceCache) + number(optimizer?.stats?.reusedRecent) +
      number(optimizer?.stats?.reusedInFlight) + number(coalescer?.stats?.duplicateCallsShared);
    const durationMs = Math.max(1, nowMs() - new Date(value.startedAt).getTime());
    const readsPerMinute = (estimatedMin / durationMs) * 60000;
    const projected = Math.round((estimatedMin / durationMs) * 86400000);
    const threshold = number(localStorage.getItem('varga_fs_diag_daily_threshold')) || DEFAULT_DAILY_THRESHOLD;
    const progress = Math.min(100, (estimatedMin / threshold) * 100);
    const issues = recommendations(value, legacy);
    const high = issues.filter((item) => item.severity === 'high').length;
    const medium = issues.filter((item) => item.severity === 'medium').length;
    const status = high > 0 ? 'high' : medium > 0 ? 'medium' : 'ok';
    const statusLabel = status === 'high' ? 'ATTENZIONE' : status === 'medium' ? 'DA CONTROLLARE' : 'REGOLARE';
    const minuteItems = Object.entries(value.minuteBuckets || {}).sort((a, b) => a[0].localeCompare(b[0])).slice(-12);
    const minuteMax = Math.max(1, ...minuteItems.map(([, stats]) => number(stats?.estimatedReadsMin)));
    const timeline = minuteItems.length ? minuteItems.map(([minute, stats]) => {
      const count = number(stats?.estimatedReadsMin);
      const height = Math.max(5, Math.round((count / minuteMax) * 100));
      return `<span style="height:${height}%" data-label="${escapeHtml(`${minute.slice(11)} · ${formatNumber(count)} letture`)}"></span>`;
    }).join('') : '<span style="height:5%" data-label="Dati in raccolta"></span>';

    root.innerHTML = `
      <div class="fs-v4-head"><div><h3>📊 Quadro immediato Firestore V4</h3><div class="fs-v4-note">Distingue payload, cache, modifiche reali e stima delle letture lato server.</div></div><div class="fs-v4-actions"><button class="btn" type="button" data-v4-threshold>SOGLIA</button><button class="btn" type="button" data-v4-reset>AZZERA TUTTO</button><button class="btn" type="button" data-v4-export>SCARICA REPORT COMPLETO</button></div></div>
      <div class="fs-v4-status"><span class="fs-v4-light ${status}"></span><div><strong>${statusLabel}</strong><small>${high ? `${high} criticità importanti` : medium ? `${medium} aspetti da verificare` : 'Nessuna criticità evidente'} · monitor attivo da ${Math.max(1, Math.round(durationMs / 60000))} min</small></div></div>
      <div class="fs-v4-kpis">
        <div class="fs-v4-kpi"><span>Letture server stimate</span><strong>${formatNumber(estimatedMin)}</strong><small>intervallo ${formatNumber(estimatedMin)}–${formatNumber(estimatedMax)}</small></div>
        <div class="fs-v4-kpi"><span>Documenti dalla cache</span><strong>${formatNumber(cacheDocs)}</strong><small>${cachePct}% del payload noto</small></div>
        <div class="fs-v4-kpi"><span>Payload ricevuto</span><strong>${formatNumber(payloadDocs)}</strong><small>non equivale alla fatturazione</small></div>
        <div class="fs-v4-kpi"><span>Query evitate</span><strong>${formatNumber(avoided)}</strong><small>cache e richieste condivise</small></div>
        <div class="fs-v4-kpi"><span>Listener fisici attivi</span><strong>${active.length}</strong><small>${duplicateGroups.length} gruppi duplicati</small></div>
        <div class="fs-v4-kpi"><span>Ritmo corrente</span><strong>${formatNumber(readsPerMinute)}/min</strong><small>proiezione prova: ${formatNumber(projected)}/24h</small></div>
      </div>
      <div class="fs-v4-grid">
        <div class="fs-v4-panel"><h4>Origine dei documenti ricevuti</h4><div class="fs-v4-donut-wrap"><div class="fs-v4-donut" style="background:conic-gradient(#2563eb 0 ${serverPct}%,#16a34a ${serverPct}% ${serverPct + cachePct}%,#94a3b8 ${serverPct + cachePct}% 100%)"></div><div class="fs-v4-legend"><div><span class="fs-v4-dot server"></span>Server: ${formatNumber(serverDocs)}</div><div><span class="fs-v4-dot cache"></span>Cache: ${formatNumber(cacheDocs)}</div><div><span class="fs-v4-dot unknown"></span>Non noto: ${formatNumber(unknownDocs)}</div></div></div></div>
        <div class="fs-v4-panel"><h4>Collezioni più costose — stima server</h4>${renderBars(value.collections)}</div>
        <div class="fs-v4-panel"><h4>Andamento ultime 12 finestre/minuti</h4><div class="fs-v4-timeline">${timeline}</div></div>
        <div class="fs-v4-panel"><h4>Soglia giornaliera configurata: ${formatNumber(threshold)}</h4><strong>${formatNumber(estimatedMin)} raccolte oggi dalla V4</strong><div class="fs-v4-progress"><span style="width:${progress}%"></span></div><p class="fs-v4-note">La proiezione usa soltanto la velocità della sessione corrente e non sostituisce i dati ufficiali Google Cloud.</p></div>
      </div>
      <div class="fs-v4-alerts">${issues.slice(0, 8).map((item) => `<div class="fs-v4-alert ${item.severity}">${escapeHtml(item.message)}</div>`).join('')}</div>`;

    root.querySelector('[data-v4-export]')?.addEventListener('click', downloadReport);
    root.querySelector('[data-v4-threshold]')?.addEventListener('click', () => {
      const requested = prompt('Soglia giornaliera di riferimento per il semaforo:', String(threshold));
      if (requested === null) return;
      const parsed = number(requested);
      if (parsed > 0) {
        localStorage.setItem('varga_fs_diag_daily_threshold', String(parsed));
        renderDashboard();
      }
    });
    root.querySelector('[data-v4-reset]')?.addEventListener('click', () => {
      if (!confirm('Azzerare tutta la diagnostica Firestore V3 e V4 di oggi?')) return;
      try { window.VargaFirestoreDiagnostics?.reset?.(); } catch (_) {}
      localStorage.setItem(storageKey(), JSON.stringify(blank()));
      location.reload();
    });
    const legacyReset = card.querySelector('[data-reset]');
    if (legacyReset && legacyReset.dataset.v4ResetHook !== 'true') {
      legacyReset.dataset.v4ResetHook = 'true';
      legacyReset.onclick = () => {
        if (!confirm('Azzerare tutta la diagnostica Firestore V3 e V4 di oggi?')) return;
        try { window.VargaFirestoreDiagnostics?.reset?.(); } catch (_) {}
        localStorage.setItem(storageKey(), JSON.stringify(blank()));
        location.reload();
      };
    }
    return true;
  }

  function recordLifecycle(type, extra = {}) {
    const value = load();
    value.lifecycle.unshift({ at: nowIso(), type, screen: currentScreen(), url: location.href, online: navigator.onLine, visibilityState: document.visibilityState, ...extra });
    persist(value);
  }

  function installLifecycle() {
    if (lifecycleInstalled) return;
    lifecycleInstalled = true;
    recordLifecycle('diagnostics-v4-start');
    document.addEventListener('visibilitychange', () => recordLifecycle('visibility-change'));
    window.addEventListener('online', () => recordLifecycle('network-online'));
    window.addEventListener('offline', () => recordLifecycle('network-offline'));
    window.addEventListener('hashchange', () => recordLifecycle('hash-change'));
    window.addEventListener('popstate', () => recordLifecycle('history-popstate'));
    window.addEventListener('pagehide', (event) => recordLifecycle('page-hide', { persisted: Boolean(event.persisted) }));
    const connection = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
    connection?.addEventListener?.('change', () => recordLifecycle('connection-change', {
      effectiveType: connection.effectiveType,
      downlinkMbps: connection.downlink,
      rttMs: connection.rtt,
      saveData: connection.saveData
    }));
  }

  function init() {
    installLifecycle();
    if (!patch()) {
      let attempts = 0;
      const timer = setInterval(() => {
        attempts += 1;
        if (patch() || attempts >= 80) clearInterval(timer);
      }, 100);
    }
    let renderAttempts = 0;
    const renderInterval = setInterval(() => {
      renderAttempts += 1;
      if (renderDashboard() || renderAttempts >= 80) clearInterval(renderInterval);
    }, 250);
    const controlRoot = document.getElementById('control-center-content') || document.getElementById('control-center-page');
    if (controlRoot) new MutationObserver(() => setTimeout(renderDashboard, 0)).observe(controlRoot, { childList: true, subtree: false });
  }

  window.VargaFirestoreDiagnosticsV4 = {
    installed: true,
    version: SCRIPT_VERSION,
    read: load,
    buildReport,
    download: downloadReport,
    render: renderDashboard,
    reset() {
      localStorage.setItem(storageKey(), JSON.stringify(blank()));
      renderDashboard();
    },
    setDailyThreshold(value) {
      localStorage.setItem('varga_fs_diag_daily_threshold', String(Math.max(1, number(value))));
      renderDashboard();
    },
    activeListeners() {
      return Object.values(load().listeners || {}).filter((item) => item.active);
    }
  };

  // Tenta subito: deve avvolgere Firestore prima del multiplexer dei listener.
  patch();
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();

(() => {
  'use strict';

  if (window.__vargaFirestoreSafeOptimizer) return;
  window.__vargaFirestoreSafeOptimizer = true;

  const pendingReads = new Map();
  const sharedListeners = new Map();
  const MAX_PENDING_AGE_MS = 30000;

  function stableValue(value) {
    if (value == null) return '';
    if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return String(value);
    try {
      return JSON.stringify(value, Object.keys(value).sort());
    } catch {
      return String(value);
    }
  }

  function canonicalQuery(query) {
    const internal = query?._query;
    return internal?.canonicalId?.()
      || internal?.path?.canonicalString?.()
      || internal?.path?.toString?.()
      || query?.path
      || '';
  }

  function queryKey(query, args) {
    return `query:${canonicalQuery(query)}:${args.map(stableValue).join('|')}`;
  }

  function documentKey(reference, args) {
    return `doc:${reference?.path || ''}:${args.map(stableValue).join('|')}`;
  }

  function coalesce(key, execute) {
    const now = Date.now();
    const existing = pendingReads.get(key);
    if (existing && now - existing.startedAt < MAX_PENDING_AGE_MS) return existing.promise;

    const result = execute();
    if (!result || typeof result.then !== 'function') return result;

    const tracked = Promise.resolve(result);
    pendingReads.set(key, { promise: tracked, startedAt: now });
    tracked.finally(() => {
      const current = pendingReads.get(key);
      if (current?.promise === tracked) pendingReads.delete(key);
    }).catch(() => {});
    return tracked;
  }

  function patchReadMethod(prototype, method, keyBuilder) {
    if (!prototype || typeof prototype[method] !== 'function') return false;
    const current = prototype[method];
    if (current.__vargaSafeReadOptimizer) return true;

    const wrapped = function safeCoalescedFirestoreRead(...args) {
      return coalesce(keyBuilder(this, args), () => current.apply(this, args));
    };
    wrapped.__vargaSafeReadOptimizer = true;
    wrapped.__vargaOriginal = current;
    prototype[method] = wrapped;
    return true;
  }

  function normalizeSnapshotArguments(args) {
    let options = null;
    let observer = null;
    let next = null;
    let error = null;
    let complete = null;
    let offset = 0;

    if (args[0] && typeof args[0] === 'object' && typeof args[0] !== 'function'
      && ('includeMetadataChanges' in args[0])) {
      options = args[0];
      offset = 1;
    }

    const handler = args[offset];
    if (handler && typeof handler === 'object') {
      observer = handler;
      next = typeof handler.next === 'function' ? handler.next.bind(handler) : null;
      error = typeof handler.error === 'function' ? handler.error.bind(handler) : null;
      complete = typeof handler.complete === 'function' ? handler.complete.bind(handler) : null;
    } else {
      next = typeof handler === 'function' ? handler : null;
      error = typeof args[offset + 1] === 'function' ? args[offset + 1] : null;
      complete = typeof args[offset + 2] === 'function' ? args[offset + 2] : null;
    }

    return { options, observer, next, error, complete };
  }

  function listenerKey(query, parsed) {
    const metadata = parsed.options?.includeMetadataChanges === true ? 'metadata:1' : 'metadata:0';
    return `listener:${canonicalQuery(query)}:${metadata}`;
  }

  function notify(entry, type, value) {
    for (const subscriber of Array.from(entry.subscribers)) {
      const callback = subscriber[type];
      if (typeof callback !== 'function') continue;
      try {
        callback(value);
      } catch (callbackError) {
        window.setTimeout(() => { throw callbackError; }, 0);
      }
    }
  }

  function patchQueryListeners(prototype) {
    if (!prototype || typeof prototype.onSnapshot !== 'function') return false;
    const current = prototype.onSnapshot;
    if (current.__vargaSharedListenerOptimizer) return true;

    const wrapped = function sharedFirestoreListener(...args) {
      const parsed = normalizeSnapshotArguments(args);
      if (!parsed.next && !parsed.error && !parsed.complete) return current.apply(this, args);

      const key = listenerKey(this, parsed);
      let entry = sharedListeners.get(key);
      const subscriber = {
        next: parsed.next,
        error: parsed.error,
        complete: parsed.complete,
        active: true
      };

      if (!entry) {
        entry = { key, subscribers: new Set(), unsubscribe: null };
        sharedListeners.set(key, entry);
        entry.subscribers.add(subscriber);

        const onNext = (snapshot) => notify(entry, 'next', snapshot);
        const onError = (listenerError) => {
          notify(entry, 'error', listenerError);
          sharedListeners.delete(key);
          entry.subscribers.clear();
        };
        const onComplete = () => {
          notify(entry, 'complete');
          sharedListeners.delete(key);
          entry.subscribers.clear();
        };

        try {
          entry.unsubscribe = parsed.options
            ? current.call(this, parsed.options, onNext, onError, onComplete)
            : current.call(this, onNext, onError, onComplete);
        } catch (listenerError) {
          sharedListeners.delete(key);
          entry.subscribers.clear();
          throw listenerError;
        }
      } else {
        entry.subscribers.add(subscriber);
      }

      return function unsubscribeSharedFirestoreListener() {
        if (!subscriber.active) return;
        subscriber.active = false;
        entry.subscribers.delete(subscriber);
        if (entry.subscribers.size > 0) return;
        sharedListeners.delete(key);
        try {
          entry.unsubscribe?.();
        } catch (error) {
          console.warn('Chiusura listener Firestore condiviso non riuscita:', error);
        }
      };
    };

    wrapped.__vargaSharedListenerOptimizer = true;
    wrapped.__vargaOriginal = current;
    prototype.onSnapshot = wrapped;
    return true;
  }

  function patchFirestore() {
    const firestore = window.firebase?.firestore;
    if (!firestore) return false;

    const queryReads = patchReadMethod(firestore.Query?.prototype, 'get', queryKey);
    const documentReads = patchReadMethod(firestore.DocumentReference?.prototype, 'get', documentKey);
    const queryListeners = patchQueryListeners(firestore.Query?.prototype);

    if (queryReads || documentReads || queryListeners) {
      console.info('Ottimizzazione Firestore attiva: letture simultanee e listener identici vengono condivisi mantenendo il realtime.');
      return true;
    }
    return false;
  }

  if (!patchFirestore()) {
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      if (patchFirestore() || attempts >= 60) window.clearInterval(timer);
    }, 250);
  }

  window.VargaFirestoreSafeOptimizer = {
    pendingCount: () => pendingReads.size,
    sharedListenerCount: () => sharedListeners.size,
    sharedSubscriberCount: () => Array.from(sharedListeners.values()).reduce((sum, entry) => sum + entry.subscribers.size, 0),
    clearPendingReads: () => pendingReads.clear()
  };
})();

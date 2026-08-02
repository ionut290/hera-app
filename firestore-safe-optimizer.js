(() => {
  'use strict';

  if (window.__vargaFirestoreSafeOptimizer) return;
  window.__vargaFirestoreSafeOptimizer = true;

  const pendingReads = new Map();
  const MAX_PENDING_AGE_MS = 30000;

  function stableValue(value) {
    if (value == null) return '';
    if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
      return String(value);
    }
    try {
      return JSON.stringify(value, Object.keys(value).sort());
    } catch {
      return String(value);
    }
  }

  function queryKey(query, args) {
    const internal = query?._query;
    const canonical =
      internal?.canonicalId?.() ||
      internal?.path?.canonicalString?.() ||
      internal?.path?.toString?.() ||
      query?.path ||
      '';
    return `query:${canonical}:${args.map(stableValue).join('|')}`;
  }

  function documentKey(reference, args) {
    return `doc:${reference?.path || ''}:${args.map(stableValue).join('|')}`;
  }

  function coalesce(key, execute) {
    const now = Date.now();
    const existing = pendingReads.get(key);
    if (existing && now - existing.startedAt < MAX_PENDING_AGE_MS) return existing.promise;

    let result;
    try {
      result = execute();
    } catch (error) {
      throw error;
    }

    if (!result || typeof result.then !== 'function') return result;

    const tracked = Promise.resolve(result);
    pendingReads.set(key, { promise: tracked, startedAt: now });
    tracked.finally(() => {
      const current = pendingReads.get(key);
      if (current?.promise === tracked) pendingReads.delete(key);
    }).catch(() => {});
    return tracked;
  }

  function patchMethod(prototype, method, keyBuilder) {
    if (!prototype || typeof prototype[method] !== 'function') return false;
    const current = prototype[method];
    if (current.__vargaSafeReadOptimizer) return true;

    const wrapped = function safeCoalescedFirestoreRead(...args) {
      const key = keyBuilder(this, args);
      return coalesce(key, () => current.apply(this, args));
    };
    wrapped.__vargaSafeReadOptimizer = true;
    wrapped.__vargaOriginal = current;
    prototype[method] = wrapped;
    return true;
  }

  function patchFirestore() {
    const firestore = window.firebase?.firestore;
    if (!firestore) return false;

    const queryPatched = patchMethod(firestore.Query?.prototype, 'get', queryKey);
    const documentPatched = patchMethod(firestore.DocumentReference?.prototype, 'get', documentKey);

    if (queryPatched || documentPatched) {
      console.info('Ottimizzazione Firestore attiva: richieste identiche simultanee condividono la stessa lettura.');
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
    clear: () => pendingReads.clear()
  };
})();

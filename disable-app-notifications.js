(() => {
  'use strict';

  if (window.__vargaAppNotificationsDisabled) return;
  window.__vargaAppNotificationsDisabled = true;

  const state = {
    installed: false,
    attempts: 0,
    blockedListeners: 0,
    blockedReads: 0,
    blockedWrites: 0,
    collection: 'appNotifications'
  };
  window.VargaAppNotificationsDisabled = state;

  const bindOrReturn = (target, property) => {
    const value = Reflect.get(target, property, target);
    return typeof value === 'function' ? value.bind(target) : value;
  };

  const emptyQuerySnapshot = () => ({
    docs: [],
    size: 0,
    empty: true,
    metadata: { fromCache: true, hasPendingWrites: false },
    forEach() {},
    docChanges() { return []; }
  });

  const emptyDocumentSnapshot = (ref) => ({
    id: ref?.id || '',
    ref,
    exists: false,
    metadata: { fromCache: true, hasPendingWrites: false },
    data() { return undefined; },
    get() { return undefined; }
  });

  function deliverEmptySnapshot(args, snapshotFactory) {
    const observer = args.find((value) => value && typeof value === 'object' && typeof value.next === 'function');
    const next = observer?.next || args.find((value) => typeof value === 'function');
    if (typeof next !== 'function') return;
    Promise.resolve().then(() => {
      try {
        next(snapshotFactory());
      } catch (error) {
        console.warn('Callback appNotifications disattivato ha generato un errore:', error);
      }
    });
  }

  function wrapDocument(reference) {
    return new Proxy(reference, {
      get(target, property) {
        if (property === 'get') {
          return async () => {
            state.blockedReads += 1;
            return emptyDocumentSnapshot(target);
          };
        }
        if (property === 'onSnapshot') {
          return (...args) => {
            state.blockedListeners += 1;
            deliverEmptySnapshot(args, () => emptyDocumentSnapshot(target));
            return () => {};
          };
        }
        if (property === 'set' || property === 'update' || property === 'delete') {
          return async () => {
            state.blockedWrites += 1;
            return undefined;
          };
        }
        return bindOrReturn(target, property);
      }
    });
  }

  function wrapQuery(query) {
    return new Proxy(query, {
      get(target, property) {
        if (property === 'get') {
          return async () => {
            state.blockedReads += 1;
            return emptyQuerySnapshot();
          };
        }
        if (property === 'onSnapshot') {
          return (...args) => {
            state.blockedListeners += 1;
            deliverEmptySnapshot(args, emptyQuerySnapshot);
            return () => {};
          };
        }
        if (['where', 'orderBy', 'limit', 'limitToLast', 'startAt', 'startAfter', 'endAt', 'endBefore'].includes(property)) {
          return (...args) => wrapQuery(target[property](...args));
        }
        return bindOrReturn(target, property);
      }
    });
  }

  function wrapCollection(reference) {
    return new Proxy(reference, {
      get(target, property) {
        if (property === 'add') {
          return async () => {
            state.blockedWrites += 1;
            return {
              id: 'appNotifications-disabled',
              path: 'appNotifications/appNotifications-disabled'
            };
          };
        }
        if (property === 'doc') {
          return (...args) => wrapDocument(target.doc(...args));
        }
        if (property === 'get') {
          return async () => {
            state.blockedReads += 1;
            return emptyQuerySnapshot();
          };
        }
        if (property === 'onSnapshot') {
          return (...args) => {
            state.blockedListeners += 1;
            deliverEmptySnapshot(args, emptyQuerySnapshot);
            return () => {};
          };
        }
        if (['where', 'orderBy', 'limit', 'limitToLast', 'startAt', 'startAfter', 'endAt', 'endBefore'].includes(property)) {
          return (...args) => wrapQuery(target[property](...args));
        }
        return bindOrReturn(target, property);
      }
    });
  }

  function install() {
    state.attempts += 1;
    const FirestoreCtor = window.firebase?.firestore?.Firestore;
    const prototype = FirestoreCtor?.prototype;
    if (!prototype || typeof prototype.collection !== 'function') return false;

    if (prototype.collection.__vargaAppNotificationsDisabled) {
      state.installed = true;
      return true;
    }

    const originalCollection = prototype.collection;
    const disabledCollection = function disabledAppNotificationsCollection(path) {
      const reference = originalCollection.call(this, path);
      return String(path) === state.collection ? wrapCollection(reference) : reference;
    };

    Object.defineProperty(disabledCollection, '__vargaAppNotificationsDisabled', {
      value: true,
      configurable: false,
      enumerable: false,
      writable: false
    });

    try {
      prototype.collection = disabledCollection;
    } catch (error) {
      console.warn('Impossibile disattivare appNotifications:', error);
      return false;
    }

    state.installed = prototype.collection === disabledCollection;
    if (state.installed) {
      console.info('appNotifications DISATTIVATO: nessuna lettura, listener o nuova scrittura.');
    }
    return state.installed;
  }

  if (install()) return;

  const timer = setInterval(() => {
    if (install() || state.attempts >= 200) {
      clearInterval(timer);
      if (!state.installed) console.warn('Disattivazione appNotifications non installata.');
    }
  }, 25);
})();

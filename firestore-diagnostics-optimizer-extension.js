(() => {
  'use strict';

  if (window.__vargaFsOptimizerDiagnosticsExtension) return;
  window.__vargaFsOptimizerDiagnosticsExtension = true;

  const text = (value) => String(value ?? '').trim();
  const number = (value) => Number.isFinite(Number(value)) ? Number(value) : 0;

  function readOptimizerState() {
    const optimizer = window.HeraFirestoreRegistryOptimizer;
    const capturedAt = new Date().toISOString();
    if (!optimizer || typeof optimizer.getState !== 'function') {
      return {
        capturedAt,
        installed: Boolean(optimizer?.installed),
        available: false,
        stats: {},
        inFlight: 0,
        recent: 0,
        profileWriteGuards: 0,
        note: 'Ottimizzatore non disponibile in questa sessione.'
      };
    }

    try {
      const state = optimizer.getState() || {};
      return {
        capturedAt,
        installed: Boolean(optimizer.installed),
        available: true,
        stats: { ...(state.stats || optimizer.stats || {}) },
        inFlight: number(state.inFlight),
        recent: number(state.recent),
        profileWriteGuards: number(state.profileWriteGuards),
        note: 'Contatori della sessione corrente; si azzerano quando l’app viene chiusa e riaperta.'
      };
    } catch (error) {
      return {
        capturedAt,
        installed: Boolean(optimizer.installed),
        available: false,
        stats: {},
        inFlight: 0,
        recent: 0,
        profileWriteGuards: 0,
        error: text(error?.message || error),
        note: 'Stato dell’ottimizzatore non leggibile.'
      };
    }
  }

  function readInflightCoalescerState() {
    const coalescer = window.HeraFirestoreInflightReadCoalescer;
    const capturedAt = new Date().toISOString();
    if (!coalescer || typeof coalescer.getState !== 'function') {
      return {
        capturedAt,
        installed: Boolean(coalescer?.installed),
        available: false,
        inFlight: 0,
        stats: {},
        note: 'Condivisione richieste in corso non disponibile in questa sessione.'
      };
    }
    try {
      const state = coalescer.getState() || {};
      return {
        capturedAt,
        installed: Boolean(state.installed ?? coalescer.installed),
        available: true,
        inFlight: number(state.inFlight),
        stats: { ...(state.stats || coalescer.stats || {}) },
        note: 'Condivide solo richieste simultanee identiche; non conserva snapshot dopo il completamento.'
      };
    } catch (error) {
      return {
        capturedAt,
        installed: Boolean(coalescer.installed),
        available: false,
        inFlight: 0,
        stats: {},
        error: text(error?.message || error),
        note: 'Stato della condivisione richieste non leggibile.'
      };
    }
  }

  function enhancedReport() {
    const diagnostics = window.VargaFirestoreDiagnostics;
    const base = typeof diagnostics?.read === 'function' ? diagnostics.read() : {};
    return {
      ...base,
      activeListeners: typeof diagnostics?.activeListeners === 'function'
        ? diagnostics.activeListeners()
        : number(base.activeListeners),
      registryOptimizer: readOptimizerState(),
      inflightReadCoalescer: readInflightCoalescerState()
    };
  }

  function downloadEnhancedReport() {
    const value = enhancedReport();
    const date = value.date || new Date().toLocaleDateString('en-CA', { timeZone: 'Europe/Rome' });
    const url = URL.createObjectURL(new Blob([JSON.stringify(value, null, 2)], {
      type: 'application/json'
    }));
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `diagnostica-firestore-${date}.json`;
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function row(label, value) {
    return `<div class="control-center-row"><span>${label}</span><strong>${value}</strong></div>`;
  }

  function renderOptimizerSummary() {
    const card = document.getElementById('firestore-operation-diagnostics-card');
    const summary = card?.querySelector('[data-summary]');
    if (!summary) return false;

    let panel = summary.querySelector('[data-registry-optimizer-summary]');
    if (!panel) {
      panel = document.createElement('details');
      panel.dataset.registryOptimizerSummary = 'true';
      panel.open = true;
      summary.prepend(panel);
    }

    const state = readOptimizerState();
    const stats = state.stats || {};
    const coalescerState = readInflightCoalescerState();
    const coalescerStats = coalescerState.stats || {};
    const avoided = number(stats.reusedDeviceCache)
      + number(stats.reusedRecent)
      + number(stats.reusedInFlight)
      + number(coalescerStats.duplicateCallsShared);

    panel.innerHTML = `
      <summary><strong>Ottimizzazione letture — sessione corrente</strong></summary>
      ${row('Ottimizzatore disponibile', state.available ? 'SÌ' : 'NO')}
      ${row('Query Firestore di rete', number(stats.networkGets))}
      ${row('Risposte dalla cache del dispositivo', number(stats.reusedDeviceCache))}
      ${row('Risposte dalla cache recente', number(stats.reusedRecent))}
      ${row('Richieste duplicate cache condivise', number(stats.reusedInFlight))}
      ${row('Coalescer richieste in corso disponibile', coalescerState.available ? 'SÌ' : 'NO')}
      ${row('Letture originali avviate dal coalescer', number(coalescerStats.networkRequestsStarted))}
      ${row('Richieste simultanee condivise in sicurezza', number(coalescerStats.duplicateCallsShared))}
      ${row('Richieste attualmente in corso', number(coalescerState.inFlight))}
      ${row('Query evitate complessivamente', avoided)}
      ${row('Aggiornamenti della cache locale', number(stats.deviceCacheWrites))}
      ${row('Snapshot ricevuti dai listener', number(stats.listenerSnapshots))}
      ${row('Invalidazioni cache', number(stats.invalidations))}
      ${row('Scritture profilo lasciate passare', number(stats.profileWritesPassed))}
      ${row('Scritture profilo duplicate evitate', number(stats.profileWritesSkipped))}
      <p class="muted">${coalescerState.note}</p>`;

    if (!summary.__vargaOptimizerSummaryObserver) {
      const observer = new MutationObserver(() => {
        if (!summary.querySelector('[data-registry-optimizer-summary]')) {
          setTimeout(renderOptimizerSummary, 0);
        }
      });
      observer.observe(summary, { childList: true });
      summary.__vargaOptimizerSummaryObserver = observer;
    }
    return true;
  }

  document.addEventListener('click', (event) => {
    const exportButton = event.target.closest?.('#firestore-operation-diagnostics-card [data-export]');
    if (exportButton) {
      event.preventDefault();
      event.stopImmediatePropagation();
      downloadEnhancedReport();
      return;
    }

    if (event.target.closest?.('#firestore-operation-diagnostics-card [data-refresh], #firestore-operation-diagnostics-card [data-reset]')) {
      setTimeout(renderOptimizerSummary, 0);
    }
  }, true);

  function init() {
    let attempts = 0;
    const interval = setInterval(() => {
      attempts += 1;
      if (renderOptimizerSummary() || attempts >= 40) clearInterval(interval);
    }, 250);
  }

  window.VargaFirestoreOptimizerDiagnostics = {
    read: enhancedReport,
    optimizerState: readOptimizerState,
    inflightCoalescerState: readInflightCoalescerState,
    render: renderOptimizerSummary,
    download: downloadEnhancedReport
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();

(() => {
  'use strict';

  if (window.__vargaAppNotificationsReadOptimizer) return;
  window.__vargaAppNotificationsReadOptimizer = true;

  const state = {
    installed: false,
    attempts: 0,
    calls: 0,
    rewrittenLimits: 0,
    fallbacks: 0,
    originalLimit: 40,
    optimizedLimit: 1
  };
  window.VargaAppNotificationsReadOptimizer = state;

  function bindOrReturn(target, property) {
    const value = Reflect.get(target, property, target);
    return typeof value === 'function' ? value.bind(target) : value;
  }

  function wrapNotificationQuery(query) {
    return new Proxy(query, {
      get(target, property) {
        if (property === 'limit') {
          return (amount) => {
            const requested = Number(amount);
            if (requested === state.originalLimit) {
              state.rewrittenLimits += 1;
              return target.limit(state.optimizedLimit);
            }
            return target.limit(amount);
          };
        }
        return bindOrReturn(target, property);
      }
    });
  }

  function wrapNotificationsCollection(collectionReference) {
    return new Proxy(collectionReference, {
      get(target, property) {
        if (property === 'orderBy') {
          return (...args) => wrapNotificationQuery(target.orderBy(...args));
        }
        return bindOrReturn(target, property);
      }
    });
  }

  function install() {
    state.attempts += 1;
    if (typeof subscribeGlobalNotifications !== 'function') return false;
    if (typeof db === 'undefined' || !db || typeof db.collection !== 'function') return false;
    if (subscribeGlobalNotifications.__vargaAppNotificationsLimitOne) {
      state.installed = true;
      return true;
    }

    const originalSubscribe = subscribeGlobalNotifications;
    const optimizedSubscribe = async function optimizedSubscribeGlobalNotifications(...args) {
      const originalCollection = db.collection;
      const temporaryCollection = function optimizedCollection(path) {
        const reference = originalCollection.call(this, path);
        return String(path) === 'appNotifications'
          ? wrapNotificationsCollection(reference)
          : reference;
      };

      let overrideApplied = false;
      try {
        db.collection = temporaryCollection;
        overrideApplied = db.collection === temporaryCollection;
      } catch (error) {
        console.warn('Ottimizzazione appNotifications non applicabile: uso il listener originale.', error);
      }

      if (!overrideApplied) {
        state.fallbacks += 1;
        return originalSubscribe.apply(this, args);
      }

      state.calls += 1;
      try {
        return await originalSubscribe.apply(this, args);
      } finally {
        if (db.collection === temporaryCollection) {
          db.collection = originalCollection;
        }
      }
    };

    Object.defineProperty(optimizedSubscribe, '__vargaAppNotificationsLimitOne', {
      value: true,
      configurable: false,
      enumerable: false,
      writable: false
    });

    subscribeGlobalNotifications = optimizedSubscribe;
    state.installed = true;
    console.info('Ottimizzazione appNotifications attiva: listener iniziale limitato a 1 documento.');
    return true;
  }

  if (install()) return;

  const timer = setInterval(() => {
    if (install() || state.attempts >= 200) {
      clearInterval(timer);
      if (!state.installed) {
        console.warn('Ottimizzazione appNotifications non installata: funzione non disponibile.');
      }
    }
  }, 50);
})();

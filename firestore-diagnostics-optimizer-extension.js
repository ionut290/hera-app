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

  function enhancedReport() {
    const diagnostics = window.VargaFirestoreDiagnostics;
    const base = typeof diagnostics?.read === 'function' ? diagnostics.read() : {};
    return {
      ...base,
      activeListeners: typeof diagnostics?.activeListeners === 'function'
        ? diagnostics.activeListeners()
        : number(base.activeListeners),
      registryOptimizer: readOptimizerState()
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
    const avoided = number(stats.reusedDeviceCache)
      + number(stats.reusedRecent)
      + number(stats.reusedInFlight);

    panel.innerHTML = `
      <summary><strong>Cache personale e mezzi — sessione corrente</strong></summary>
      ${row('Ottimizzatore disponibile', state.available ? 'SÌ' : 'NO')}
      ${row('Query Firestore di rete', number(stats.networkGets))}
      ${row('Risposte dalla cache del dispositivo', number(stats.reusedDeviceCache))}
      ${row('Risposte dalla cache recente', number(stats.reusedRecent))}
      ${row('Richieste duplicate condivise', number(stats.reusedInFlight))}
      ${row('Query evitate complessivamente', avoided)}
      ${row('Aggiornamenti della cache locale', number(stats.deviceCacheWrites))}
      ${row('Snapshot ricevuti dai listener', number(stats.listenerSnapshots))}
      ${row('Invalidazioni cache', number(stats.invalidations))}
      ${row('Scritture profilo lasciate passare', number(stats.profileWritesPassed))}
      ${row('Scritture profilo duplicate evitate', number(stats.profileWritesSkipped))}
      <p class="muted">${state.note}</p>`;

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
    render: renderOptimizerSummary,
    download: downloadEnhancedReport
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();

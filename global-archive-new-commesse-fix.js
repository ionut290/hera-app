(() => {
  'use strict';

  let started = false;

  function loadPreventiviSelectionPersistence() {
    if (document.querySelector('script[data-preventivi-selection-persistence]')) return;
    const waitForModules = () => {
      if (!window.HeraPreventivi || !window.HeraPreventiviModels || !document.querySelector('script[data-preventivi-module="clients"]')) {
        window.setTimeout(waitForModules, 400);
        return;
      }
      if (document.querySelector('script[data-preventivi-selection-persistence]')) return;
      const script = document.createElement('script');
      script.src = './preventivi-persistenza-selezioni-fix.js?v=20260801a';
      script.defer = true;
      script.dataset.preventiviSelectionPersistence = '1';
      script.addEventListener('load', loadCompleteQuotePersistence, { once: true });
      script.addEventListener('error', () => {
        console.warn('Persistenza selezioni preventivi non caricata.');
        loadCompleteQuotePersistence();
      }, { once: true });
      document.head.appendChild(script);
    };
    waitForModules();
  }

  function loadCompleteQuotePersistence() {
    if (document.querySelector('script[data-preventivi-complete-persistence-v2]')) return;
    const wait = () => {
      if (!window.HeraPreventivi?.saveRemote || !window.HeraPreventiviModels?.selectedModel) {
        window.setTimeout(wait, 350);
        return;
      }
      if (document.querySelector('script[data-preventivi-complete-persistence-v2]')) return;
      const script = document.createElement('script');
      script.src = './preventivi-salvataggio-completo-v2.js?v=20260801a';
      script.defer = true;
      script.dataset.preventiviCompletePersistenceV2 = '1';
      script.addEventListener('load', loadTemporaryArchive, { once: true });
      script.addEventListener('error', () => {
        console.warn('Salvataggio completo preventivi v2 non caricato.');
        loadTemporaryArchive();
      }, { once: true });
      document.head.appendChild(script);
    };
    wait();
  }

  function loadTemporaryArchive() {
    if (document.querySelector('script[data-preventivi-temporary-archive]')) return;
    const wait = () => {
      if (!window.HeraPreventivi?.saveRemote || !window.HeraPreventivi?.persistLocal) {
        window.setTimeout(wait, 350);
        return;
      }
      if (document.querySelector('script[data-preventivi-temporary-archive]')) return;
      const script = document.createElement('script');
      script.src = './preventivi-archivio-temporaneo-30gg.js?v=20260801a';
      script.defer = true;
      script.dataset.preventiviTemporaryArchive = '1';
      script.addEventListener('error', () => console.warn('Archivio temporaneo di 30 giorni non caricato.'), { once: true });
      document.head.appendChild(script);
    };
    wait();
  }

  function start() {
    if (started) return;
    started = true;
    loadPreventiviSelectionPersistence();
    window.setTimeout(loadCompleteQuotePersistence, 2500);
    window.setTimeout(loadTemporaryArchive, 4000);
  }

  window.HeraLoadPreventiviCompatibilityFixes = start;
  window.addEventListener('hera:preventivi-ready', start, { once: true });

  if (window.HeraPreventivi && window.HeraPreventiviModels) start();
})();

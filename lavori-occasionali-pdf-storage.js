(() => {
  'use strict';

  if (window.HeraCantiereDocumentsLoaderInstalled) return;
  window.HeraCantiereDocumentsLoaderInstalled = true;

  const VERSION = '20260823-docs5';

  function ensureStyle() {
    if (document.querySelector('link[data-cantiere-doc-style]')) return;
    const style = document.createElement('link');
    style.rel = 'stylesheet';
    style.href = `documentazione-cantiere.css?v=${VERSION}`;
    style.dataset.cantiereDocStyle = '1';
    document.head.appendChild(style);
  }

  function loadScript({ selector, src, datasetKey }) {
    return new Promise((resolve) => {
      const existing = document.querySelector(selector);
      if (existing) {
        existing.addEventListener('load', resolve, { once: true });
        window.setTimeout(resolve, 1200);
        return;
      }
      const script = document.createElement('script');
      script.src = src;
      script.async = false;
      if (datasetKey) script.dataset[datasetKey] = '1';
      script.onload = resolve;
      script.onerror = () => {
        console.error(`Caricamento modulo non riuscito: ${src}`);
        resolve();
      };
      document.head.appendChild(script);
    });
  }

  async function init() {
    ensureStyle();
    if (!window.HeraCantiereDocuments?.installed) {
      await loadScript({
        selector: 'script[data-cantiere-doc-script]',
        src: `documentazione-cantiere.js?v=${VERSION}`,
        datasetKey: 'cantiereDocScript'
      });
    }
    if (!window.HeraOccasionalDocumentsRuntime?.installed) {
      await loadScript({
        selector: 'script[data-occasional-doc-runtime]',
        src: `documentazione-occasionali-runtime.js?v=${VERSION}`,
        datasetKey: 'occasionalDocRuntime'
      });
    }
    if (!window.HeraAllPlantsDocumentsRuntime?.installed) {
      await loadScript({
        selector: 'script[data-all-plants-doc-runtime]',
        src: `documentazione-tutte-commesse-runtime.js?v=${VERSION}`,
        datasetKey: 'allPlantsDocRuntime'
      });
    }
    window.HeraOccasionalDocumentsRuntime?.refresh?.();
    window.HeraAllPlantsDocumentsRuntime?.refresh?.();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else void init();
})();
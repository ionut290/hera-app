(() => {
  'use strict';

  if (window.__heraPreventiviLazyLoaderInstalled) return;
  window.__heraPreventiviLazyLoaderInstalled = true;

  const MENU_ID = 'open-preventivi-btn';
  let loadPromise = null;

  const modules = [
    ['core', './preventivi-core.js?v=20260801-drive1'],
    ['storage', './preventivi-storage-config.js?v=20260801-cost1'],
    ['paths', './preventivi-firestore-path-fix.js?v=20260731a'],
    ['chunks', './preventivi-firestore-chunks.js?v=20260731a'],
    ['batch-size', './preventivi-firestore-batch-fix.js?v=20260731a'],
    ['price-lists', './preventivi-price-lists.js?v=20260731a'],
    ['quotes', './preventivi-quotes.js?v=20260731a'],
    ['consuntivi', './preventivi-consuntivi.js?v=20260801b'],
    ['models-core', './preventivi-models-core.js?v=20260801b'],
    ['models-ui', './preventivi-models-ui.js?v=20260801b'],
    ['models-documents', './preventivi-models-documents.js?v=20260801d'],
    ['models-export', './preventivi-models-export.js?v=20260801c'],
    ['conditional-discount', './preventivi-ribasso-condizionale-fix.js?v=20260801a'],
    ['registry-model-export-fix', './preventivi-registry-model-export-fix.js?v=20260801a'],
    ['registry-model-followup', './preventivi-registry-model-followup.js?v=20260801b'],
    ['registry-fix', './preventivi-commesse-impianti-fix.js?v=20260801c'],
    ['tabs-models-plant-guard', './preventivi-tabs-models-plant-guard.js?v=20260801a'],
    ['matrix-runtime-fix', './preventivi-matrix-runtime-fix.js?v=20260801a'],
    ['draft-preserver', './preventivi-draft-preserver.js?v=20260801b'],
    ['commessa-search-bridge', './preventivi-commessa-search-bridge.js?v=20260801a'],
    ['clients', './preventivi-clienti-feature.js?v=20260801-cost1'],
    ['feature', './preventivi-feature.js?v=20260731a']
  ];

  function ensureMenuButton() {
    let button = document.getElementById(MENU_ID);
    if (button) return button;
    const section = document.getElementById('menu-operativo-title')?.closest('.menu-section');
    if (!section) return null;
    button = document.createElement('button');
    button.id = MENU_ID;
    button.className = 'btn menu-title-btn';
    button.type = 'button';
    button.innerHTML = '<span class="menu-item-icon" aria-hidden="true">🧾</span><span data-preventivi-menu-label>Preventivi</span>';
    const anchor = document.getElementById('open-segnalazioni-btn');
    if (anchor?.parentElement === section) anchor.insertAdjacentElement('afterend', button);
    else section.appendChild(button);
    return button;
  }

  function setButtonState(state) {
    const button = ensureMenuButton();
    if (!button) return;
    const label = button.querySelector('[data-preventivi-menu-label]') || button;
    if (state === 'loading') {
      button.disabled = true;
      button.setAttribute('aria-busy', 'true');
      label.textContent = 'Caricamento…';
    } else {
      button.disabled = false;
      button.removeAttribute('aria-busy');
      label.textContent = 'Preventivi';
    }
  }

  function ensureStyles() {
    if (!document.querySelector('link[data-preventivi-feature-css]')) {
      const css = document.createElement('link');
      css.rel = 'stylesheet';
      css.href = './preventivi-feature.css?v=20260731a';
      css.dataset.preventiviFeatureCss = '1';
      document.head.appendChild(css);
    }
    if (!document.querySelector('link[data-preventivi-models-css]')) {
      const css = document.createElement('link');
      css.rel = 'stylesheet';
      css.href = './preventivi-models.css?v=20260801a';
      css.dataset.preventiviModelsCss = '1';
      document.head.appendChild(css);
    }
  }

  function loadScript(name, src) {
    return new Promise((resolve, reject) => {
      const selector = `script[data-preventivi-module="${name}"]`;
      const existing = document.querySelector(selector);
      if (existing?.dataset.loaded === '1') return resolve();
      if (existing) {
        existing.addEventListener('load', resolve, { once: true });
        existing.addEventListener('error', reject, { once: true });
        return;
      }
      const script = document.createElement('script');
      script.src = src;
      script.dataset.preventiviModule = name;
      script.addEventListener('load', () => {
        script.dataset.loaded = '1';
        resolve();
      }, { once: true });
      script.addEventListener('error', () => reject(new Error(`Modulo Preventivi non caricato: ${name}`)), { once: true });
      document.head.appendChild(script);
    });
  }

  function loadAll() {
    if (window.HeraPreventivi?.open && document.querySelector('script[data-preventivi-module="feature"][data-loaded="1"]')) {
      return Promise.resolve();
    }
    if (loadPromise) return loadPromise;
    ensureStyles();
    setButtonState('loading');
    loadPromise = modules.reduce((promise, [name, src]) => promise.then(() => loadScript(name, src)), Promise.resolve())
      .then(() => {
        window.dispatchEvent(new CustomEvent('hera:preventivi-ready'));
        window.HeraLoadPreventiviCompatibilityFixes?.();
      })
      .catch((error) => {
        console.error('Caricamento Preventivi non riuscito:', error);
        loadPromise = null;
        window.alert('Preventivi non caricati. Controlla la connessione e riprova.');
        throw error;
      })
      .finally(() => setButtonState('ready'));
    return loadPromise;
  }

  async function openPreventivi() {
    try {
      await loadAll();
      window.HeraPreventivi?.open?.();
    } catch (_) {
      // Il messaggio è già stato mostrato da loadAll.
    }
  }

  document.addEventListener('click', (event) => {
    const button = event.target?.closest?.(`#${MENU_ID}`);
    if (!button) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    void openPreventivi();
  }, true);

  window.HeraPreventiviLazyLoader = Object.freeze({ load: loadAll, open: openPreventivi });

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', ensureMenuButton, { once: true });
  else ensureMenuButton();
})();

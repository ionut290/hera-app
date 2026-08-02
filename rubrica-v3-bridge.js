(() => {
  'use strict';
  if (window.__vargaRubricaV3Bridge) return;
  window.__vargaRubricaV3Bridge = true;

  let replacing = false;

  function openCloudRubrica() {
    if (replacing || document.querySelector('.rcv3')) return;
    if (typeof window.openRubricaCloudV3 !== 'function') return;
    replacing = true;
    try {
      document.getElementById('rubrica-feature-v2-page')?.remove();
      document.getElementById('rubrica-feature-page')?.remove();
      window.openRubricaCloudV3();
    } finally {
      window.setTimeout(() => { replacing = false; }, 200);
    }
  }

  function replaceLegacyPage() {
    const legacy = document.getElementById('rubrica-feature-v2-page')
      || document.getElementById('rubrica-feature-page');
    if (!legacy || document.querySelector('.rcv3')) return;
    openCloudRubrica();
  }

  function installObserver() {
    if (!document.body || window.__vargaRubricaV3Observer) return;
    const observer = new MutationObserver(() => replaceLegacyPage());
    observer.observe(document.body, { childList: true, subtree: true });
    window.__vargaRubricaV3Observer = observer;
    replaceLegacyPage();
  }

  document.addEventListener('click', (event) => {
    const target = event.target?.closest?.('button,a,[role="button"]');
    if (!target || target.closest('.rcv3')) return;
    const label = String(target.textContent || '').trim().toUpperCase();
    if (label === 'RUBRICA' || label.includes('RUBRICA CONTATTI')) {
      window.setTimeout(openCloudRubrica, 0);
      window.setTimeout(openCloudRubrica, 150);
    }
  }, true);

  function init() {
    installObserver();
    let attempts = 0;
    const timer = window.setInterval(() => {
      attempts += 1;
      replaceLegacyPage();
      if (typeof window.openRubricaCloudV3 === 'function' || attempts >= 80) {
        window.clearInterval(timer);
      }
    }, 250);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();

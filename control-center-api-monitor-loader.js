(() => {
  'use strict';

  const VERSION = '20260823-live3';
  if (window.HeraControlCenterApiMonitorLoader?.version === VERSION) return;

  const MONITOR_SRC = `./control-center-firestore-usage.js?v=${VERSION}`;
  const MONITOR_PATH = '/control-center-firestore-usage.js';
  const FIXED_IDS = ['street-view-usage-control-card', 'api-usage-control-card'];
  let observer = null;

  function assetPath(value) {
    try { return new URL(String(value || ''), document.baseURI).pathname; }
    catch (_) { return String(value || '').split('?')[0].split('#')[0]; }
  }

  function pinCards() {
    const root = document.getElementById('control-center-content');
    const firestoreCard = document.getElementById('firestore-usage-control-card');
    if (!root || !firestoreCard) return;

    let anchor = firestoreCard;
    FIXED_IDS.forEach((id) => {
      const card = document.getElementById(id);
      if (!card) return;
      card.classList.remove('control-center-card', 'cc-source-hidden');
      card.classList.add('cc-fixed-api-monitor-card');
      if (card.parentElement !== root || card.previousElementSibling !== anchor) {
        anchor.insertAdjacentElement('afterend', card);
      }
      anchor = card;
    });
  }

  function watchCards() {
    if (observer) return;
    observer = new MutationObserver(() => pinCards());
    observer.observe(document.documentElement, { childList: true, subtree: true });
    window.addEventListener('hashchange', () => {
      window.setTimeout(pinCards, 0);
      window.setTimeout(pinCards, 300);
      window.setTimeout(pinCards, 900);
    });
    document.addEventListener('click', (event) => {
      if (!event.target.closest('#open-control-center-btn')) return;
      window.setTimeout(pinCards, 0);
      window.setTimeout(pinCards, 300);
      window.setTimeout(pinCards, 900);
    }, true);
  }

  function loadLatestMonitor() {
    const existing = Array.from(document.scripts || []).find((script) => assetPath(script.src) === MONITOR_PATH);
    if (existing && existing.src.includes(`v=${VERSION}`)) {
      watchCards();
      pinCards();
      return;
    }

    // Questo loader viene eseguito da firebase-config prima del vecchio loader del menu.
    // Una versione precedente eventualmente rimasta in cache non deve impedire l'upgrade.
    if (window.__heraFirestoreUsageControlInstalled) {
      try { delete window.__heraFirestoreUsageControlInstalled; }
      catch (_) { window.__heraFirestoreUsageControlInstalled = false; }
    }

    const script = document.createElement('script');
    script.src = MONITOR_SRC;
    script.dataset.firestoreUsageControlLatest = VERSION;
    script.addEventListener('load', () => {
      watchCards();
      pinCards();
      window.setTimeout(pinCards, 300);
    }, { once: true });
    script.addEventListener('error', () => console.warn('Monitor API Centro di controllo non caricato.'), { once: true });
    document.head.appendChild(script);
  }

  watchCards();
  loadLatestMonitor();

  window.HeraControlCenterApiMonitorLoader = {
    installed: true,
    version: VERSION,
    refresh: pinCards
  };
})();

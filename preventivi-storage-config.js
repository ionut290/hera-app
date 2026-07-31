(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  if (!PV) throw new Error('Preventivi core non caricato.');

  // Percorsi protetti dalle regole Firestore della commessa.
  PV.collections.priceLists = 'commesse/__preventivi__/prezziari';
  PV.collections.quotes = 'commesse/__preventivi__/preventivi';

  const STORAGE_MODE_KEY = 'hera_preventivi_storage_mode_v1';
  const FIRESTORE_MODE = 'firestore';
  const DEVICE_MODE = 'device';

  const readStorageMode = () => {
    try {
      return localStorage.getItem(STORAGE_MODE_KEY) === DEVICE_MODE ? DEVICE_MODE : FIRESTORE_MODE;
    } catch (_) {
      return FIRESTORE_MODE;
    }
  };

  const writeStorageMode = (mode) => {
    try { localStorage.setItem(STORAGE_MODE_KEY, mode); }
    catch (error) { console.warn('Preventivi: preferenza di salvataggio non memorizzata.', error); }
  };

  PV.state.storageMode = readStorageMode();
  PV.state.firestoreConnecting = false;
  PV.state.storageAuthUnsubscribe = null;

  const originalSetSyncBadge = PV.setSyncBadge.bind(PV);

  function installStorageStyles() {
    if (document.getElementById('preventivi-storage-mode-style')) return;
    const style = document.createElement('style');
    style.id = 'preventivi-storage-mode-style';
    style.textContent = `
      .pv-storage-button{border:0;cursor:pointer;font:inherit;text-align:left}
      .pv-storage-button:hover{filter:brightness(1.12)}
      .pv-storage-button:focus-visible{outline:3px solid rgba(96,165,250,.55);outline-offset:2px}
      .pv-storage-modal{position:fixed;inset:0;z-index:2600;display:grid;place-items:center;padding:16px;background:rgba(15,23,42,.66);backdrop-filter:blur(5px)}
      .pv-storage-modal.hidden{display:none!important}
      .pv-storage-card{width:min(560px,100%);max-height:90vh;overflow:auto;padding:20px;border:1px solid #dbe3ec;border-radius:20px;background:#fff;color:#172033;box-shadow:0 28px 80px rgba(15,23,42,.35)}
      .pv-storage-card h2{margin:0 0 7px;font-size:1.25rem}
      .pv-storage-card>p{margin:0 0 16px;color:#667085;line-height:1.45}
      .pv-storage-options{display:grid;gap:10px}
      .pv-storage-option{display:grid;grid-template-columns:42px minmax(0,1fr);gap:11px;align-items:center;width:100%;padding:13px;border:1px solid #cfd8e3;border-radius:14px;background:#f8fafc;color:#172033;text-align:left;cursor:pointer}
      .pv-storage-option:hover{border-color:#60a5fa;background:#eff6ff}
      .pv-storage-option.is-active{border-color:#16a34a;background:#f0fdf4;box-shadow:inset 0 0 0 1px rgba(22,163,74,.12)}
      .pv-storage-option-icon{display:grid;place-items:center;width:42px;height:42px;border-radius:12px;background:#fff;font-size:1.35rem}
      .pv-storage-option strong,.pv-storage-option small{display:block}
      .pv-storage-option small{margin-top:3px;color:#667085;line-height:1.35}
      .pv-storage-actions{display:flex;justify-content:flex-end;margin-top:15px}
      .pv-storage-close{min-height:42px;padding:8px 14px;border:1px solid #cfd8e3;border-radius:10px;background:#fff;font:inherit;font-weight:750;cursor:pointer}
      @media(max-width:720px){.pv-header .pv-storage-button{display:inline-flex!important;max-width:48px;width:48px;height:40px;justify-content:center;padding:7px;overflow:hidden;color:transparent}.pv-header .pv-storage-button::first-letter{color:initial}.pv-storage-card{padding:16px;border-radius:18px}}
    `;
    document.head.appendChild(style);
  }

  function storageButton() {
    return PV.page()?.querySelector('[data-pv-sync]') || null;
  }

  function storageStatusLabel(message = '') {
    if (PV.state.storageMode === DEVICE_MODE) return '💾 Solo dispositivo';
    if (PV.state.remoteDenied) return '⚠️ Firestore non autorizzato';
    if (PV.state.remoteConnected) return '☁️ Firestore sincronizzato';
    const user = window.firebase?.auth?.()?.currentUser;
    if (!user) return '☁️ Firestore: accedi';
    if (/salvato sul dispositivo|modalità dispositivo/i.test(message)) return '☁️ Firestore: dati in attesa';
    return '☁️ Firestore: connessione…';
  }

  PV.setSyncBadge = (message = '', type = '') => {
    originalSetSyncBadge(message, type);
    const button = storageButton();
    if (!button) return;
    const label = storageStatusLabel(message);
    button.textContent = label;
    button.title = `${label}. Premi per scegliere dove salvare i dati.`;
    button.setAttribute('aria-label', button.title);
    button.dataset.state = PV.state.storageMode === DEVICE_MODE
      ? 'warning'
      : PV.state.remoteConnected ? 'ok' : PV.state.remoteDenied ? 'warning' : type;
  };

  function ensureStorageModal() {
    let modal = document.getElementById('pv-storage-modal');
    if (modal) return modal;
    modal = document.createElement('section');
    modal.id = 'pv-storage-modal';
    modal.className = 'pv-storage-modal hidden';
    modal.setAttribute('aria-hidden', 'true');
    modal.innerHTML = `
      <div class="pv-storage-card" role="dialog" aria-modal="true" aria-labelledby="pv-storage-title">
        <h2 id="pv-storage-title">Dove vuoi salvare i dati?</h2>
        <p>Puoi sincronizzare prezziari e preventivi tramite Firestore oppure conservarli soltanto su questo dispositivo.</p>
        <div class="pv-storage-options">
          <button type="button" class="pv-storage-option" data-pv-storage-choice="firestore">
            <span class="pv-storage-option-icon" aria-hidden="true">☁️</span>
            <span><strong>Firestore – consigliato</strong><small>I dati vengono sincronizzati e sono disponibili sugli altri dispositivi autorizzati.</small></span>
          </button>
          <button type="button" class="pv-storage-option" data-pv-storage-choice="device">
            <span class="pv-storage-option-icon" aria-hidden="true">💾</span>
            <span><strong>Solo dispositivo</strong><small>I dati restano esclusivamente nel browser utilizzato e non vengono condivisi.</small></span>
          </button>
        </div>
        <div class="pv-storage-actions"><button type="button" class="pv-storage-close" data-pv-storage-close>Chiudi</button></div>
      </div>`;
    document.body.appendChild(modal);
    return modal;
  }

  function refreshStorageChoices() {
    const modal = ensureStorageModal();
    modal.querySelectorAll('[data-pv-storage-choice]').forEach((button) => {
      button.classList.toggle('is-active', button.dataset.pvStorageChoice === PV.state.storageMode);
    });
  }

  function openStorageModal() {
    refreshStorageChoices();
    const modal = ensureStorageModal();
    modal.classList.remove('hidden');
    modal.setAttribute('aria-hidden', 'false');
  }

  function closeStorageModal() {
    const modal = document.getElementById('pv-storage-modal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.setAttribute('aria-hidden', 'true');
  }

  function ensureStorageUi() {
    installStorageStyles();
    const current = storageButton();
    if (!current) return;
    let button = current;
    if (current.tagName !== 'BUTTON') {
      button = document.createElement('button');
      button.type = 'button';
      button.className = current.className;
      [...current.attributes].forEach((attribute) => {
        if (attribute.name !== 'class') button.setAttribute(attribute.name, attribute.value);
      });
      current.replaceWith(button);
    }
    button.type = 'button';
    button.classList.add('pv-storage-button');
    button.dataset.pvStorageButton = '1';
    PV.setSyncBadge('', PV.state.storageMode === DEVICE_MODE ? 'warning' : '');
  }

  const originalEnsurePage = PV.ensurePage.bind(PV);
  PV.ensurePage = () => {
    originalEnsurePage();
    ensureStorageUi();
  };

  function disconnectFirestoreSubscriptions() {
    (PV.state.unsubscribers || []).forEach((unsubscribe) => {
      try { unsubscribe(); } catch (_) { /* Listener già chiuso. */ }
    });
    PV.state.unsubscribers = [];
    window.clearTimeout(PV.state.syncTimer);
    PV.state.syncTimer = null;
    PV.state.firestore = null;
    PV.state.remoteConnected = false;
    PV.state.remoteDenied = false;
    PV.state.firestoreConnecting = false;
  }

  function ensureAuthWatcher() {
    if (PV.state.storageAuthUnsubscribe || !window.firebase?.auth) return;
    try {
      PV.state.storageAuthUnsubscribe = window.firebase.auth().onAuthStateChanged((user) => {
        if (PV.state.storageMode !== FIRESTORE_MODE) return;
        disconnectFirestoreSubscriptions();
        if (user) PV.connectFirebase();
        else PV.setSyncBadge('Accesso richiesto per Firestore.', 'warning');
      });
    } catch (error) {
      console.warn('Preventivi: controllo login Firestore non disponibile.', error);
    }
  }

  PV.connectFirebase = () => {
    ensureAuthWatcher();
    if (PV.state.storageMode !== FIRESTORE_MODE) {
      disconnectFirestoreSubscriptions();
      PV.setSyncBadge('Salvataggio locale selezionato.', 'warning');
      return;
    }
    if (!window.firebase?.firestore || !window.firebase?.auth) {
      PV.setSyncBadge('Firestore non disponibile.', 'warning');
      return;
    }
    const user = window.firebase.auth().currentUser;
    if (!user) {
      PV.setSyncBadge('Accesso richiesto per Firestore.', 'warning');
      return;
    }
    if (PV.state.firestoreConnecting || (PV.state.firestore && PV.state.unsubscribers.length)) return;

    try {
      PV.state.firestoreConnecting = true;
      PV.state.remoteConnected = false;
      PV.state.remoteDenied = false;
      PV.state.firestore = window.firebase.firestore();
      PV.setSyncBadge('Connessione a Firestore…');
      PV.subscribeCollection(PV.collections.priceLists, 'priceLists', 'priceLists');
      PV.subscribeCollection(PV.collections.quotes, 'quotes', 'quotes');
      PV.state.firestoreConnecting = false;
      PV.scheduleSync();
    } catch (error) {
      PV.state.firestoreConnecting = false;
      PV.state.firestore = null;
      console.warn('Preventivi: Firestore non disponibile.', error);
      PV.setSyncBadge('Firestore non disponibile.', 'warning');
    }
  };

  PV.setStorageMode = (mode) => {
    const normalized = mode === DEVICE_MODE ? DEVICE_MODE : FIRESTORE_MODE;
    PV.state.storageMode = normalized;
    writeStorageMode(normalized);
    disconnectFirestoreSubscriptions();

    if (normalized === FIRESTORE_MODE) {
      // Porta su Firestore anche gli eventuali dati creati in precedenza sul dispositivo.
      PV.state.priceLists = (PV.state.priceLists || []).map((item) => ({ ...item, syncPending: true }));
      PV.state.quotes = (PV.state.quotes || []).map((item) => ({ ...item, syncPending: true }));
      PV.persistLocal?.();
      PV.setSyncBadge('Connessione a Firestore…');
      PV.connectFirebase();
    } else {
      PV.setSyncBadge('Salvataggio locale selezionato.', 'warning');
    }
    refreshStorageChoices();
    closeStorageModal();
    PV.renderCurrentView?.();
  };

  document.addEventListener('click', (event) => {
    if (event.target.closest('[data-pv-storage-button]')) {
      event.preventDefault();
      event.stopPropagation();
      openStorageModal();
      return;
    }
    const choice = event.target.closest('[data-pv-storage-choice]');
    if (choice) {
      event.preventDefault();
      PV.setStorageMode(choice.dataset.pvStorageChoice);
      return;
    }
    const modal = event.target.closest('#pv-storage-modal');
    if (event.target.closest('[data-pv-storage-close]') || event.target.id === 'pv-storage-modal') {
      event.preventDefault();
      closeStorageModal();
    } else if (!modal) {
      return;
    }
  }, true);

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') closeStorageModal();
  });

  installStorageStyles();
  ensureAuthWatcher();

  // Carica in modo isolato la matrice Excel e l'aggiornamento completo dei prezziari.
  if (!document.querySelector('link[data-preventivi-matrix-css]')) {
    const css = document.createElement('link');
    css.rel = 'stylesheet';
    css.href = './preventivi-price-list-matrix.css?v=20260731b';
    css.dataset.preventiviMatrixCss = '1';
    document.head.appendChild(css);
  }

  if (!document.querySelector('script[data-preventivi-matrix-js]')) {
    const script = document.createElement('script');
    script.src = './preventivi-price-list-matrix.js?v=20260731b';
    script.dataset.preventiviMatrixJs = '1';
    script.addEventListener('error', () => {
      console.warn('Preventivi: modulo matrice prezziario non caricato.');
    }, { once: true });
    document.head.appendChild(script);
  }
})();

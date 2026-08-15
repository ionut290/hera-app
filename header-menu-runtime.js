(() => {
  'use strict';

  if (window.__heraCommessaListenerCleanupInstalled) return;

  const originalSubscribeImpianti = window.subscribeImpianti;
  const originalStopImpiantiSubscription = window.stopImpiantiSubscription;
  const originalSubscribeCommessaNotes = window.subscribeCommessaNotes;
  const originalStopCommessaNotesSubscription = window.stopCommessaNotesSubscription;

  if (typeof originalSubscribeImpianti !== 'function' ||
      typeof originalSubscribeCommessaNotes !== 'function') {
    console.warn('Pulizia listener commessa non installata: funzioni principali non disponibili.');
  } else {
    let impiantiSubscriptionStarted = false;
    let notesSubscriptionStarted = false;

    window.subscribeImpianti = function subscribeImpiantiWithCleanup(...args) {
      if (impiantiSubscriptionStarted && typeof originalStopImpiantiSubscription === 'function') {
        originalStopImpiantiSubscription();
      }

      const result = originalSubscribeImpianti.apply(this, args);
      impiantiSubscriptionStarted = true;
      return result;
    };

    window.subscribeCommessaNotes = function subscribeCommessaNotesWithCleanup(...args) {
      if (notesSubscriptionStarted && typeof originalStopCommessaNotesSubscription === 'function') {
        originalStopCommessaNotesSubscription();
      }

      const result = originalSubscribeCommessaNotes.apply(this, args);
      notesSubscriptionStarted = true;
      return result;
    };

    window.__heraCommessaListenerCleanupInstalled = true;
    console.info('Pulizia listener commessa installata.');
  }
})();

(() => {
  const qs = (selector) => document.querySelector(selector);

  function createLabel(text) {
    const label = document.createElement('span');
    label.className = 'header-action-label';
    label.textContent = text;
    return label;
  }

  function normalizeActionButton(button, labelText, iconText) {
    if (!button) return;
    button.classList.add('header-action-button');
    button.textContent = '';
    const icon = document.createElement('span');
    icon.className = 'header-action-icon';
    icon.setAttribute('aria-hidden', 'true');
    icon.textContent = iconText;
    button.append(icon, createLabel(labelText));
  }

  function setupProfileButton() {
    const button = qs('#user-toggle-btn');
    if (!button) return;
    button.classList.add('header-action-button');
    button.textContent = '';
    const image = document.createElement('img');
    image.className = 'header-login-photo';
    image.alt = 'Foto profilo Google';
    const icon = document.createElement('span');
    icon.className = 'header-action-icon';
    icon.setAttribute('aria-hidden', 'true');
    icon.textContent = '👤';
    button.append(image, icon, createLabel('Login'));
  }

  function hideHomeTitle() {
    const title = qs('#home-page .logo-head > h1');
    if (title) {
      title.classList.add('home-brand-title-hidden');
      title.setAttribute('aria-hidden', 'true');
    }
  }

  function setupHomeHeader() {
    const card = qs('#home-page .logo-card');
    if (!card) return;
    card.classList.add('home-header-modern');
    hideHomeTitle();
    setupProfileButton();
    normalizeActionButton(qs('#refresh-app-btn'), 'Refresh', '↻');
    normalizeActionButton(qs('#update-app-btn'), 'Aggiorna', '↻');
    normalizeActionButton(qs('#menu-toggle-btn'), 'Menu', '☰');
    normalizeActionButton(qs('#snow-service-btn'), 'Neve', '❄️');
    const notificationButton = qs('#notification-inbox-btn, #notification-bell-btn, #notification-center-btn, .notification-bell-btn, [data-notification-bell]');
    if (notificationButton) {
      notificationButton.classList.add('header-action-button', 'header-notification-button');
      notificationButton.querySelector('[aria-hidden="true"]')?.classList.add('header-action-icon');
      if (!notificationButton.querySelector('.header-action-label')) notificationButton.append(createLabel('Notifiche'));
    }
  }

  function setupSnowHeader() {
    normalizeActionButton(qs('#snow-refresh-app-btn'), 'Refresh', '↻');
    normalizeActionButton(qs('#snow-service-menu-btn'), 'Menu', '☰');
  }

  function loadFirestorePresenceCostGuard() {
    try {
      if (document.querySelector('script[data-firestore-presence-cost-guard]')) return;
      const script = document.createElement('script');
      script.src = './firestore-presence-cost-guard.js?v=20260802a';
      script.defer = true;
      script.dataset.firestorePresenceCostGuard = '1';
      script.addEventListener('error', () => console.warn('Riduzione scritture presenza Firestore non caricata.'), { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Riduzione scritture presenza Firestore non caricato:', error);
    }
  }

  function loadBrandingFeature() {
    try {
      if (document.querySelector('script[data-varga-branding]')) return;
      const script = document.createElement('script');
      script.src = './varga-branding.js?v=20260804-diagnostics-reset1';
      script.dataset.vargaBranding = '1';
      script.addEventListener('error', () => console.warn('Branding VARGA CANTIERI non caricato.'), { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Branding VARGA CANTIERI non caricato:', error);
    }
  }

  function setLoginPhoto() {
    const button = qs('#user-toggle-btn');
    if (!button) return;
    const image = button.querySelector('.header-login-photo');
    const user = window.firebase?.auth?.()?.currentUser || null;
    const photoURL = user?.photoURL || '';
    if (image && photoURL) {
      image.src = photoURL;
      image.alt = user?.displayName ? `Foto profilo di ${user.displayName}` : 'Foto profilo Google';
      button.classList.add('has-photo');
    } else {
      button.classList.remove('has-photo');
      if (image) image.removeAttribute('src');
    }
  }

  function loadOperatorProfileFeature() {
    try {
      if (!document.querySelector('link[data-operator-profile-css]')) {
        const css = document.createElement('link');
        css.rel = 'stylesheet';
        css.href = './operator-profile-feature.css?v=20260731-safe2';
        css.dataset.operatorProfileCss = '1';
        document.head.appendChild(css);
      }
      if (!document.querySelector('script[data-operator-profile-js]')) {
        const script = document.createElement('script');
        script.src = './operator-profile-feature.js?v=20260731-safe2';
        script.defer = true;
        script.dataset.operatorProfileJs = '1';
        script.addEventListener('error', () => console.warn('Scheda operatore non caricata: file JavaScript non disponibile.'), { once: true });
        document.head.appendChild(script);
      }
    } catch (error) {
      console.warn('Scheda operatore non caricato:', error);
    }
  }

  function loadGlobalArchiveFeature() {
    try {
      if (!document.querySelector('script[data-global-archive-sync]')) {
        const script = document.createElement('script');
        script.src = './global-archive-sync.js?v=20260801-cost1';
        script.defer = true;
        script.dataset.globalArchiveSync = '1';
        script.addEventListener('error', () => console.warn('Archivio Global permanente non caricato.'), { once: true });
        document.head.appendChild(script);
      }
      if (!document.querySelector('script[data-global-archive-new-commesse-fix]')) {
        const fix = document.createElement('script');
        fix.src = './global-archive-new-commesse-fix.js?v=20260801-lazy1';
        fix.defer = true;
        fix.dataset.globalArchiveNewCommesseFix = '1';
        fix.addEventListener('error', () => console.warn('Controllo nuove commesse Global non caricato.'), { once: true });
        document.head.appendChild(fix);
      }
    } catch (error) {
      console.warn('Archivio Global permanente non caricato:', error);
    }
  }

  function loadPreventiviLazyFeature() {
    try {
      if (document.querySelector('script[data-preventivi-lazy-loader]')) return;
      const script = document.createElement('script');
      script.src = './preventivi-lazy-loader.js?v=20260801a';
      script.defer = true;
      script.dataset.preventiviLazyLoader = '1';
      script.addEventListener('error', () => console.warn('Caricatore Preventivi non disponibile.'), { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Caricatore Preventivi non disponibile:', error);
    }
  }

  function loadFirestoreUsageControl() {
    try {
      if (document.querySelector('script[data-firestore-usage-control]')) return;
      const script = document.createElement('script');
      script.src = './control-center-firestore-usage.js?v=20260802a';
      script.defer = true;
      script.dataset.firestoreUsageControl = '1';
      script.addEventListener('error', () => console.warn('Monitoraggio consumo Firestore non caricato.'), { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Monitoraggio consumo Firestore non disponibile:', error);
    }
  }

  function loadControlCenterBackup() {
    try {
      if (document.querySelector('script[data-control-center-backup]')) return;
      const script = document.createElement('script');
      script.src = './control-center-backup.js?v=20260802a';
      script.defer = true;
      script.dataset.controlCenterBackup = '1';
      script.addEventListener('error', () => console.warn('Backup dati del Centro di controllo non caricato.'), { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Backup dati del Centro di controllo non caricato:', error);
    }
  }

  function loadPerformanceDiagnostic() {
    try {
      if (document.querySelector('script[data-performance-diagnostic]')) return;
      const script = document.createElement('script');
      script.src = './admin-console.js?v=20260803-performance1';
      script.defer = true;
      script.dataset.performanceDiagnostic = '1';
      script.addEventListener('error', () => console.warn('Diagnostica prestazioni e fluidità non caricata.'), { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Diagnostica prestazioni e fluidità non caricata:', error);
    }
  }

  function init() {
    loadFirestorePresenceCostGuard();
    loadBrandingFeature();
    setupHomeHeader();
    setupSnowHeader();
    setLoginPhoto();
    loadOperatorProfileFeature();
    loadGlobalArchiveFeature();
    loadPreventiviLazyFeature();
    loadFirestoreUsageControl();
    loadControlCenterBackup();
    loadPerformanceDiagnostic();
    if (window.firebase?.auth) {
      try {
        window.firebase.auth().onAuthStateChanged(setLoginPhoto);
      } catch (error) {
        console.warn('Impossibile aggiornare la foto profilo Google:', error);
      }
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();

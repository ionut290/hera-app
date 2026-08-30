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
  'use strict';

  const qs = (selector) => document.querySelector(selector);

  function normalizeAssetPath(value) {
    try {
      return new URL(String(value || ''), document.baseURI).pathname;
    } catch (_) {
      return String(value || '').split('?')[0].split('#')[0];
    }
  }

  function findExistingScript(selector, src) {
    const selected = selector ? document.querySelector(selector) : null;
    if (selected) return selected;
    const wantedPath = normalizeAssetPath(src);
    if (!wantedPath) return null;
    return Array.from(document.scripts || []).find((script) => normalizeAssetPath(script.src) === wantedPath) || null;
  }

  function loadScriptOnce({ selector, src, datasetKey, errorMessage, defer = true }) {
    try {
      const existing = findExistingScript(selector, src);
      if (existing) return existing;
      const script = document.createElement('script');
      script.src = src;
      script.defer = defer;
      if (datasetKey) script.dataset[datasetKey] = '1';
      if (errorMessage) {
        script.addEventListener('error', () => console.warn(errorMessage), { once: true });
      }
      document.head.appendChild(script);
      return script;
    } catch (error) {
      console.warn(errorMessage || `Modulo non caricato: ${src}`, error);
      return null;
    }
  }

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

  function communicationMenuSection() {
    return qs('#menu-comunicazione-title')?.closest('.menu-section') || null;
  }

  function closeMenuBeforeAction(button) {
    if (!button || button.dataset.secondaryMenuCloseBound === '1') return;
    button.dataset.secondaryMenuCloseBound = '1';
    button.addEventListener('click', () => qs('#menu-close-btn')?.click(), { capture: true });
  }

  function styleSecondaryMenuButton(button, labelText, iconText) {
    if (!button) return;
    button.classList.remove('header-action-button', 'header-notification-button', 'snow-service-btn');
    button.classList.add('btn', 'menu-title-btn', 'home-secondary-menu-action');
    button.innerHTML = '';
    const icon = document.createElement('span');
    icon.className = 'menu-item-icon';
    icon.setAttribute('aria-hidden', 'true');
    icon.textContent = iconText;
    const label = document.createElement('span');
    label.className = 'home-secondary-menu-label';
    label.textContent = labelText;
    button.append(icon, label);
    closeMenuBeforeAction(button);
  }

  function moveSecondaryActionsToMenu() {
    const section = communicationMenuSection();
    if (!section) return;

    const snowButton = qs('#snow-service-btn');
    if (snowButton && snowButton.parentElement !== section) {
      styleSecondaryMenuButton(snowButton, 'Servizio neve', '❄️');
      section.insertBefore(snowButton, section.children[1] || null);
    }

    const notificationButton = qs('#notification-inbox-btn, #notification-bell-btn, #notification-center-btn, .notification-bell-btn, [data-notification-bell]');
    if (notificationButton && notificationButton.parentElement !== section) {
      const badge = notificationButton.querySelector('.notification-bell-badge');
      styleSecondaryMenuButton(notificationButton, 'Notifiche', '🔔');
      if (badge) notificationButton.querySelector('.menu-item-icon')?.appendChild(badge);
      section.insertBefore(notificationButton, section.children[1] || null);
    }
  }

  function watchDynamicNotificationButton() {
    const headerActions = qs('#home-page .logo-head-action-icons');
    if (!headerActions || headerActions.dataset.secondaryActionsObserver === '1') return;
    headerActions.dataset.secondaryActionsObserver = '1';
    const observer = new MutationObserver(() => moveSecondaryActionsToMenu());
    observer.observe(headerActions, { childList: true });
  }

  function installFourActionHeaderStyle() {
    if (document.getElementById('home-four-action-header-style')) return;
    const style = document.createElement('style');
    style.id = 'home-four-action-header-style';
    style.textContent = `
      @media (max-width: 700px) {
        .home-header-modern .logo-head {
          grid-template-columns: repeat(4, minmax(0, 1fr)) !important;
        }
        .home-header-modern #user-toggle-btn,
        .home-header-modern #update-app-btn,
        .home-header-modern #refresh-app-btn,
        .home-header-modern #menu-toggle-btn {
          width: 100% !important;
          min-width: 0 !important;
          max-width: none !important;
          font-size: .62rem !important;
        }
        .home-header-modern .header-action-button .header-action-icon,
        .home-header-modern #user-toggle-btn .header-action-icon {
          font-size: 1.08rem !important;
        }
      }
    `;
    document.head.appendChild(style);
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
    moveSecondaryActionsToMenu();
    watchDynamicNotificationButton();
    installFourActionHeaderStyle();
  }

  function setupSnowHeader() {
    normalizeActionButton(qs('#snow-refresh-app-btn'), 'Refresh', '↻');
    normalizeActionButton(qs('#snow-service-menu-btn'), 'Menu', '☰');
  }

  function removeDeprecatedCommessaActions() {
    [
      '#map-fullscreen-btn',
      '#commessa-notes-toggle-btn',
      '#commessa-documents-btn',
      '#commessa-weather-refresh-btn',
      '#commessa-call-btn'
    ].forEach((selector) => qs(selector)?.remove());

    const weatherStatus = qs('#commessa-weather-refresh-status');
    if (weatherStatus && !weatherStatus.textContent.trim()) weatherStatus.remove();
  }

  function loadFirestorePresenceCostGuard() {
    loadScriptOnce({
      selector: 'script[data-firestore-presence-cost-guard]',
      src: './firestore-presence-cost-guard.js?v=20260802a',
      datasetKey: 'firestorePresenceCostGuard',
      errorMessage: 'Riduzione scritture presenza Firestore non caricata.'
    });
  }

  function loadBrandingFeature() {
    loadScriptOnce({
      selector: 'script[data-varga-branding]',
      src: './varga-branding.js?v=20260804-diagnostics-reset1',
      datasetKey: 'vargaBranding',
      errorMessage: 'Branding VARGA CANTIERI non caricato.'
    });
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
      loadScriptOnce({
        selector: 'script[data-operator-profile-js]',
        src: './operator-profile-feature.js?v=20260731-safe2',
        datasetKey: 'operatorProfileJs',
        errorMessage: 'Scheda operatore non caricata: file JavaScript non disponibile.'
      });
    } catch (error) {
      console.warn('Scheda operatore non caricata:', error);
    }
  }

  function loadGlobalArchiveFeature() {
    loadScriptOnce({
      selector: 'script[data-global-archive-sync]',
      src: './global-archive-sync.js?v=20260801-cost1',
      datasetKey: 'globalArchiveSync',
      errorMessage: 'Archivio Global permanente non caricato.'
    });
    loadScriptOnce({
      selector: 'script[data-global-archive-new-commesse-fix]',
      src: './global-archive-new-commesse-fix.js?v=20260801-lazy1',
      datasetKey: 'globalArchiveNewCommesseFix',
      errorMessage: 'Controllo nuove commesse Global non caricato.'
    });
  }

  function loadPreventiviLazyFeature() {
    loadScriptOnce({
      selector: 'script[data-preventivi-lazy-loader]',
      src: './preventivi-lazy-loader.js?v=20260801a',
      datasetKey: 'preventiviLazyLoader',
      errorMessage: 'Caricatore Preventivi non disponibile.'
    });
  }

  function loadFirestoreUsageControl() {
    loadScriptOnce({
      selector: 'script[data-firestore-usage-control]',
      src: './control-center-firestore-usage.js?v=20260802a',
      datasetKey: 'firestoreUsageControl',
      errorMessage: 'Monitoraggio consumo Firestore non caricato.'
    });
  }

  function loadControlCenterBackup() {
    loadScriptOnce({
      selector: 'script[data-control-center-backup]',
      src: './control-center-backup.js?v=20260802a',
      datasetKey: 'controlCenterBackup',
      errorMessage: 'Backup dati del Centro di controllo non caricato.'
    });
  }

  function loadPerformanceDiagnostic() {
    loadScriptOnce({
      selector: 'script[data-performance-diagnostic]',
      src: './admin-console.js?v=20260803-performance1',
      datasetKey: 'performanceDiagnostic',
      errorMessage: 'Diagnostica prestazioni e fluidità non caricata.'
    });
  }

  function loadModernImpiantiMap() {
    loadScriptOnce({
      selector: 'script[data-modern-impianti-map]',
      src: './map-modern-runtime.js?v=20260815a',
      datasetKey: 'modernImpiantiMap',
      errorMessage: 'Mappa moderna impianti non caricata.'
    });
    loadScriptOnce({
      selector: 'script[data-map-detail-close-fix]',
      src: './map-detail-close-fix.js?v=20260815a',
      datasetKey: 'mapDetailCloseFix',
      errorMessage: 'Correzione chiusura dettaglio impianto non caricata.'
    });
  }

  function loadMapSearchFocus() {
    loadScriptOnce({
      selector: 'script[data-map-search-focus]',
      src: './map-search-focus-runtime.js?v=20260815b',
      datasetKey: 'mapSearchFocus',
      errorMessage: 'Ricerca intelligente mappa non caricata.'
    });
  }

  function loadStreetViewCards() {
    loadScriptOnce({
      selector: 'script[data-street-view-cards]',
      src: './street-view-cards.js?v=20260830-map-gps-fix1',
      datasetKey: 'streetViewCards',
      errorMessage: 'Street View card impianti non caricato.',
      defer: false
    });
  }

  function loadUpdateCenter() {
    loadScriptOnce({
      selector: 'script[data-update-center-feature]',
      src: './update-center-feature.js?v=20260826a',
      datasetKey: 'updateCenterFeature',
      errorMessage: 'Centro aggiornamenti non caricato.'
    });
  }

  function init() {
    loadFirestorePresenceCostGuard();
    loadBrandingFeature();
    setupHomeHeader();
    setupSnowHeader();
    removeDeprecatedCommessaActions();
    setLoginPhoto();
    loadOperatorProfileFeature();
    loadGlobalArchiveFeature();
    loadPreventiviLazyFeature();
    loadFirestoreUsageControl();
    loadControlCenterBackup();
    loadPerformanceDiagnostic();
    loadModernImpiantiMap();
    loadMapSearchFocus();
    loadStreetViewCards();
    loadUpdateCenter();
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

(() => {
  'use strict';

  function installGardenerAssistanceMenuGroup() {
    const menuBody = document.querySelector('#side-menu .side-menu-body');
    if (!menuBody) return;

    let section = document.querySelector('#menu-assistenza-giardiniere-title')?.closest('.menu-section');
    if (!section) {
      section = document.createElement('section');
      section.className = 'menu-section';
      section.setAttribute('aria-labelledby', 'menu-assistenza-giardiniere-title');

      const title = document.createElement('h3');
      title.id = 'menu-assistenza-giardiniere-title';
      title.className = 'menu-section-title';
      title.textContent = 'Assistenza per il giardiniere';
      section.appendChild(title);

      const toolsTitle = document.querySelector('#menu-strumenti-title');
      const toolsSection = toolsTitle?.closest('.menu-section') || null;
      menuBody.insertBefore(section, toolsSection || null);
    }

    [
      '#open-gardening-assistant-btn',
      '#open-tree-search-btn',
      '#open-green-areas-btn',
      '#open-urban-furniture-btn',
      '#open-wastewater-plants-btn',
      '#open-equipment-assistant-btn'
    ].forEach((selector) => {
      const button = document.querySelector(selector);
      if (button && button.parentElement !== section) section.appendChild(button);
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', installGardenerAssistanceMenuGroup, { once: true });
  } else {
    installGardenerAssistanceMenuGroup();
  }
})();

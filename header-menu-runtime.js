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

  function loadBrandingFeature() {
    try {
      if (document.querySelector('script[data-varga-branding]')) return;
      const script = document.createElement('script');
      script.src = './varga-branding.js?v=20260731a';
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
      console.warn('Scheda operatore non caricata:', error);
    }
  }

  function loadGlobalArchiveFeature() {
    try {
      if (!document.querySelector('script[data-global-archive-sync]')) {
        const script = document.createElement('script');
        script.src = './global-archive-sync.js?v=20260801b';
        script.defer = true;
        script.dataset.globalArchiveSync = '1';
        script.addEventListener('error', () => console.warn('Archivio Global permanente non caricato.'), { once: true });
        document.head.appendChild(script);
      }
      if (!document.querySelector('script[data-global-archive-new-commesse-fix]')) {
        const fix = document.createElement('script');
        fix.src = './global-archive-new-commesse-fix.js?v=20260801a';
        fix.defer = true;
        fix.dataset.globalArchiveNewCommesseFix = '1';
        fix.addEventListener('error', () => console.warn('Controllo nuove commesse Global non caricato.'), { once: true });
        document.head.appendChild(fix);
      }
    } catch (error) {
      console.warn('Archivio Global permanente non caricato:', error);
    }
  }

  function loadPreventiviFeature() {
    try {
      if (!document.querySelector('link[data-preventivi-feature-css]')) {
        const css = document.createElement('link');
        css.rel = 'stylesheet';
        css.href = './preventivi-feature.css?v=20260731a';
        css.dataset.preventiviFeatureCss = '1';
        document.head.appendChild(css);
      }
      if (!document.querySelector('link[data-preventivi-models-css]')) {
        const modelsCss = document.createElement('link');
        modelsCss.rel = 'stylesheet';
        modelsCss.href = './preventivi-models.css?v=20260801a';
        modelsCss.dataset.preventiviModelsCss = '1';
        document.head.appendChild(modelsCss);
      }

      const modules = [
        ['core', './preventivi-core.js?v=20260731a'],
        ['storage', './preventivi-storage-config.js?v=20260731b'],
        ['paths', './preventivi-firestore-path-fix.js?v=20260731a'],
        ['chunks', './preventivi-firestore-chunks.js?v=20260731a'],
        ['batch-size', './preventivi-firestore-batch-fix.js?v=20260731a'],
        ['price-lists', './preventivi-price-lists.js?v=20260731a'],
        ['quotes', './preventivi-quotes.js?v=20260731a'],
        ['consuntivi', './preventivi-consuntivi.js?v=20260801a'],
        ['models-core', './preventivi-models-core.js?v=20260801b'],
        ['models-ui', './preventivi-models-ui.js?v=20260801b'],
        ['models-documents', './preventivi-models-documents.js?v=20260801b'],
        ['models-export', './preventivi-models-export.js?v=20260801b'],
        ['registry-model-export-fix', './preventivi-registry-model-export-fix.js?v=20260801a'],
        ['registry-model-followup', './preventivi-registry-model-followup.js?v=20260801a'],
        ['registry-fix', './preventivi-commesse-impianti-fix.js?v=20260801c'],
        ['tabs-models-plant-guard', './preventivi-tabs-models-plant-guard.js?v=20260801a'],
        ['matrix-runtime-fix', './preventivi-matrix-runtime-fix.js?v=20260801a'],
        ['draft-preserver', './preventivi-draft-preserver.js?v=20260801a'],
        ['feature', './preventivi-feature.js?v=20260731a']
      ];

      const loadModule = (index) => {
        if (index >= modules.length) return;
        const [name, src] = modules[index];
        const existing = document.querySelector(`script[data-preventivi-module="${name}"]`);
        if (existing) {
          if (existing.dataset.loaded === '1') loadModule(index + 1);
          else existing.addEventListener('load', () => loadModule(index + 1), { once: true });
          return;
        }
        const script = document.createElement('script');
        script.src = src;
        script.dataset.preventiviModule = name;
        script.addEventListener('load', () => {
          script.dataset.loaded = '1';
          loadModule(index + 1);
        }, { once: true });
        script.addEventListener('error', () => console.warn(`Modulo Preventivi non caricato: ${name}.`), { once: true });
        document.head.appendChild(script);
      };
      loadModule(0);
    } catch (error) {
      console.warn('Modulo Preventivi non caricato:', error);
    }
  }

  function init() {
    loadBrandingFeature();
    setupHomeHeader();
    setupSnowHeader();
    setLoginPhoto();
    loadOperatorProfileFeature();
    loadGlobalArchiveFeature();
    loadPreventiviFeature();
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

(function () {
  'use strict';

  const GESTIONALE_URL = 'https://ionut290.github.io/VARGA-GESTIONALE/';
  const DESKTOP_QUERY = '(min-width: 900px) and (pointer: fine)';

  function isDesktopBrowser() {
    const isNativeApp = Boolean(window.Capacitor?.isNativePlatform?.());
    return !isNativeApp && window.matchMedia(DESKTOP_QUERY).matches;
  }

  function wasOpenedFromGestionale() {
    return new URLSearchParams(window.location.search).get('fromGestionale') === '1';
  }

  function initialize() {
    if (!wasOpenedFromGestionale() || !isDesktopBrowser()) return;

    const bar = document.createElement('div');
    bar.id = 'gestionale-desktop-return';
    bar.setAttribute('role', 'navigation');
    bar.setAttribute('aria-label', 'Collegamento a Varga Gestionale');

    const button = document.createElement('button');
    button.type = 'button';
    button.textContent = '← TORNA AL GESTIONALE';
    button.addEventListener('click', () => window.location.assign(GESTIONALE_URL));

    bar.appendChild(button);
    document.body.appendChild(bar);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initialize, { once: true });
  } else {
    initialize();
  }
})();

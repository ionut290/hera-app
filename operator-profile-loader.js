(() => {
  'use strict';
  try {
    if (!document.querySelector('link[data-operator-profile-css]')) {
      const css = document.createElement('link');
      css.rel = 'stylesheet';
      css.href = './operator-profile-feature.css?v=20260731-safe1';
      css.dataset.operatorProfileCss = '1';
      document.head.appendChild(css);
    }
    if (!document.querySelector('script[data-operator-profile-js]')) {
      const script = document.createElement('script');
      script.src = './operator-profile-feature.js?v=20260731-safe1';
      script.defer = true;
      script.dataset.operatorProfileJs = '1';
      document.head.appendChild(script);
    }
  } catch (error) {
    console.warn('Scheda operatore non caricata:', error);
  }
})();

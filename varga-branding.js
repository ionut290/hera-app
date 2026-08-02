(() => {
  'use strict';

  const BRAND_NAME = 'VARGA CANTIERI';
  const BRAND_PATTERN = /\b(?:Hera App|Varga Cantieri)\b/gi;
  const SKIP_TAGS = new Set(['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEXTAREA', 'OPTION']);
  const BRAND_ATTRIBUTES = ['alt', 'aria-label', 'title'];

  function disableAutomaticHoursRepair(attempt = 0) {
    const current = window.repairDuplicateHours;
    if (typeof current !== 'function') {
      if (attempt < 40) window.setTimeout(() => disableAutomaticHoursRepair(attempt + 1), 50);
      return;
    }
    if (current.__vargaStartupGuard) return;
    const original = current;
    const guarded = function hoursRepairStartupGuard(options = {}) {
      if (options?.force === true) return original.apply(this, arguments);
      console.info('Riparazione automatica ore saltata: usare force:true solo per una riparazione manuale richiesta.');
      return Promise.resolve({
        changed: false,
        repaired: 0,
        deleted: 0,
        duplicates: 0,
        skipped: true,
        reason: 'automatic-startup-disabled'
      });
    };
    guarded.__vargaStartupGuard = true;
    guarded.__vargaOriginalRepair = original;
    window.repairDuplicateHours = guarded;
  }

  function replaceBrand(value) {
    return String(value ?? '').replace(BRAND_PATTERN, BRAND_NAME);
  }

  function brandTextNode(node) {
    if (!node || node.nodeType !== Node.TEXT_NODE || SKIP_TAGS.has(node.parentElement?.tagName)) return;
    const next = replaceBrand(node.nodeValue);
    if (next !== node.nodeValue) node.nodeValue = next;
  }

  function brandElement(element) {
    if (!(element instanceof Element)) return;
    BRAND_ATTRIBUTES.forEach((attribute) => {
      if (!element.hasAttribute(attribute)) return;
      const current = element.getAttribute(attribute) || '';
      const next = replaceBrand(current);
      if (next !== current) element.setAttribute(attribute, next);
    });
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
    let node = walker.nextNode();
    while (node) {
      brandTextNode(node);
      node = walker.nextNode();
    }
  }

  function applyBranding(root = document.documentElement) {
    document.title = BRAND_NAME;
    const applicationName = document.querySelector('meta[name="application-name"]');
    if (applicationName) applicationName.setAttribute('content', BRAND_NAME);
    brandElement(root);
  }

  function observeDynamicContent() {
    if (!document.body || window.__vargaBrandingObserver) return;
    const observer = new MutationObserver((mutations) => {
      mutations.forEach((mutation) => {
        if (mutation.type === 'characterData') {
          brandTextNode(mutation.target);
          return;
        }
        mutation.addedNodes.forEach((node) => {
          if (node.nodeType === Node.TEXT_NODE) brandTextNode(node);
          else if (node.nodeType === Node.ELEMENT_NODE) brandElement(node);
        });
      });
    });
    observer.observe(document.body, { childList: true, subtree: true, characterData: true });
    window.__vargaBrandingObserver = observer;
  }

  function loadFirestoreSafeOptimizer() {
    try {
      if (document.querySelector('script[data-firestore-safe-optimizer]')) return;
      const script = document.createElement('script');
      script.src = './firestore-safe-optimizer.js?v=20260802a';
      script.defer = true;
      script.dataset.firestoreSafeOptimizer = '1';
      script.addEventListener('error', () => console.warn('Ottimizzatore Firestore non caricato; avvio app non interrotto.'), { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Ottimizzatore Firestore non caricato; avvio app non interrotto:', error);
    }
  }

  function loadFirestoreDiagnostics() {
    try {
      if (document.querySelector('script[data-firestore-operation-diagnostics]')) return;
      const script = document.createElement('script');
      script.src = './firestore-operation-diagnostics.js?v=20260802b';
      script.defer = true;
      script.dataset.firestoreOperationDiagnostics = '1';
      script.addEventListener('error', () => console.warn('Diagnostica Firestore non caricata; avvio app non interrotto.'), { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Diagnostica Firestore non caricata; avvio app non interrotto:', error);
    }
  }

  function loadVcardShareFeature() {
    try {
      if (document.querySelector('script[data-rubrica-vcard-share]')) return;
      const script = document.createElement('script');
      script.src = './rubrica-vcard-share.js?v=20260802a';
      script.defer = true;
      script.dataset.rubricaVcardShare = '1';
      script.addEventListener('error', () => console.warn('Condivisione vCard Rubrica non caricata; avvio app non interrotto.'), { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Condivisione vCard Rubrica non caricata; avvio app non interrotto:', error);
    }
  }

  function loadGoogleProfileFeature() {
    try {
      if (document.querySelector('script[data-rubrica-google-profile]')) {
        loadVcardShareFeature();
        return;
      }
      const script = document.createElement('script');
      script.src = './rubrica-google-profile.js?v=20260802-google1';
      script.defer = true;
      script.dataset.rubricaGoogleProfile = '1';
      script.addEventListener('load', loadVcardShareFeature, { once: true });
      script.addEventListener('error', () => {
        console.warn('Profilo Google Rubrica non caricato; avvio app non interrotto.');
        loadVcardShareFeature();
      }, { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Profilo Google Rubrica non caricato; avvio app non interrotto:', error);
      loadVcardShareFeature();
    }
  }

  function loadRubricaFeature() {
    try {
      if (document.querySelector('script[data-rubrica-feature-v2]')) {
        loadGoogleProfileFeature();
        return;
      }

      const loadView = () => {
        if (document.querySelector('script[data-rubrica-feature-v2]')) {
          loadGoogleProfileFeature();
          return;
        }
        const script = document.createElement('script');
        script.src = './rubrica-feature-v2.js?v=20260802-email1';
        script.defer = true;
        script.dataset.rubricaFeatureV2 = '1';
        script.addEventListener('load', loadGoogleProfileFeature, { once: true });
        script.addEventListener('error', () => console.warn('Rubrica V2 non caricata; avvio app non interrotto.'), { once: true });
        document.head.appendChild(script);
      };

      if (!document.querySelector('script[data-rubrica-user-enrichment]')) {
        const enrichment = document.createElement('script');
        enrichment.src = './rubrica-user-enrichment.js?v=20260802-user1';
        enrichment.defer = true;
        enrichment.dataset.rubricaUserEnrichment = '1';
        enrichment.addEventListener('load', loadView, { once: true });
        enrichment.addEventListener('error', () => {
          console.warn('Arricchimento Rubrica non caricato; apro comunque la Rubrica base.');
          loadView();
        }, { once: true });
        document.head.appendChild(enrichment);
      } else {
        loadView();
      }
    } catch (error) {
      console.warn('Rubrica V2 non caricata; avvio app non interrotto:', error);
    }
  }

  function init() {
    disableAutomaticHoursRepair();
    loadFirestoreSafeOptimizer();
    applyBranding();
    observeDynamicContent();
    window.setTimeout(loadFirestoreDiagnostics, 0);
    window.setTimeout(loadRubricaFeature, 0);
  }

  disableAutomaticHoursRepair();
  loadFirestoreSafeOptimizer();
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();

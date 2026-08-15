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
      return Promise.resolve({ changed:false,repaired:0,deleted:0,duplicates:0,skipped:true,reason:'automatic-startup-disabled' });
    };
    guarded.__vargaStartupGuard = true;
    guarded.__vargaOriginalRepair = original;
    window.repairDuplicateHours = guarded;
  }

  function replaceBrand(value) { return String(value ?? '').replace(BRAND_PATTERN, BRAND_NAME); }
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
    while (node) { brandTextNode(node); node = walker.nextNode(); }
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
        if (mutation.type === 'characterData') { brandTextNode(mutation.target); return; }
        mutation.addedNodes.forEach((node) => {
          if (node.nodeType === Node.TEXT_NODE) brandTextNode(node);
          else if (node.nodeType === Node.ELEMENT_NODE) brandElement(node);
        });
      });
    });
    observer.observe(document.body, { childList:true, subtree:true, characterData:true });
    window.__vargaBrandingObserver = observer;
  }

  function loadScriptOnce(selector, src, datasetKey, onLoad, errorMessage) {
    try {
      const existing = document.querySelector(selector);
      if (existing) { if (onLoad) onLoad(); return existing; }
      const script = document.createElement('script');
      script.src = src;
      script.defer = true;
      script.dataset[datasetKey] = '1';
      if (onLoad) script.addEventListener('load', onLoad, { once:true });
      script.addEventListener('error', () => console.warn(errorMessage), { once:true });
      document.head.appendChild(script);
      return script;
    } catch (error) {
      console.warn(errorMessage, error);
      return null;
    }
  }

  function loadSquadValidationPersonnelCompat() {
    loadScriptOnce(
      'script[data-squad-validation-personnel-compat]',
      './squad-validation-personnel-compat.js?v=20260806a',
      'squadValidationPersonnelCompat',
      null,
      'Compatibilità controllo squadra non caricata; resta attivo il controllo originale.'
    );
  }

  function loadWhatsAppInstalledOnlyGuard() {
    loadScriptOnce(
      'script[data-whazzup-installed-only-guard]',
      './whazzup-preload-cache.js?v=20260803-installed-only2',
      'whazzupInstalledOnlyGuard',
      null,
      'Protezione WhatsApp installato non caricata; nessun fallback web deve essere usato.'
    );
  }

  function loadFirestoreDiagnostics() {
    loadScriptOnce('script[data-firestore-operation-diagnostics]', './firestore-operation-diagnostics.js?v=20260804-v3', 'firestoreOperationDiagnostics', null, 'Diagnostica Firestore non caricata; avvio app non interrotto.');
  }

  function loadVcardShareFeature() {
    loadScriptOnce('script[data-rubrica-vcard-share]', './rubrica-vcard-share.js?v=20260802a', 'rubricaVcardShare', null, 'Condivisione vCard Rubrica non caricata; avvio app non interrotto.');
  }

  function loadPersonaleRestore() {
    loadScriptOnce('script[data-rubrica-personale-restore]', './rubrica-personale-restore.js?v=20260802a', 'rubricaPersonaleRestore', null, 'Ripristino elenco personale non caricato.');
  }

  function loadRubricaV3Bridge() {
    loadScriptOnce('script[data-rubrica-v3-bridge]', './rubrica-v3-bridge.js?v=20260802-fix1', 'rubricaV3Bridge', null, 'Collegamento Rubrica V3 non caricato; resta disponibile la Rubrica base.');
  }

  function loadRubricaMatrixImport() {
    loadScriptOnce('script[data-rubrica-matrice-import]', './rubrica-matrice-personale-import.js?v=20260802a', 'rubricaMatriceImport', null, 'Importazione matrice personale non caricata.');
  }

  function loadRubricaCloudV3() {
    const afterCloud = () => { loadRubricaV3Bridge(); loadRubricaMatrixImport(); };
    loadScriptOnce('script[data-rubrica-cloud-v3]', './rubrica-cloud-v3.js?v=20260802-fix4', 'rubricaCloudV3', afterCloud, 'Rubrica condivisa V3 non caricata; resta disponibile la Rubrica base.');
  }

  function loadRubricaPermissionsBridge() {
    loadScriptOnce('script[data-rubrica-permissions-bridge]', './rubrica-firestore-permissions-bridge.js?v=20260802a', 'rubricaPermissionsBridge', loadRubricaCloudV3, 'Compatibilità permessi Rubrica non caricata.');
  }

  function loadGoogleProfileFeature() {
    const afterGoogle = () => { loadVcardShareFeature(); loadRubricaPermissionsBridge(); };
    loadScriptOnce('script[data-rubrica-google-profile]', './rubrica-google-profile.js?v=20260802-google1', 'rubricaGoogleProfile', afterGoogle, 'Profilo Google Rubrica non caricato; avvio app non interrotto.');
  }

  function loadRubricaFeature() {
    loadPersonaleRestore();
    const loadView = () => {
      loadScriptOnce('script[data-rubrica-feature-v2]', './rubrica-feature-v2.js?v=20260802-email1', 'rubricaFeatureV2', loadGoogleProfileFeature, 'Rubrica V2 non caricata; avvio app non interrotto.');
    };
    loadScriptOnce('script[data-rubrica-user-enrichment]', './rubrica-user-enrichment.js?v=20260802-user1', 'rubricaUserEnrichment', loadView, 'Arricchimento Rubrica non caricato; apro comunque la Rubrica base.');
    window.setTimeout(loadView, 1200);
    window.setTimeout(loadRubricaPermissionsBridge, 1400);
    window.setTimeout(loadRubricaCloudV3, 1800);
    window.setTimeout(loadRubricaV3Bridge, 2200);
    window.setTimeout(loadRubricaMatrixImport, 2400);
  }

  function init() {
    applyBranding();
    observeDynamicContent();
    window.setTimeout(loadFirestoreDiagnostics, 0);
    window.setTimeout(loadRubricaFeature, 0);
  }

  // I guard che devono essere disponibili il prima possibile vengono avviati una sola volta qui.
  disableAutomaticHoursRepair();
  loadSquadValidationPersonnelCompat();
  loadWhatsAppInstalledOnlyGuard();
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once:true });
  else init();
})();

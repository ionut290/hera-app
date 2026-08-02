(() => {
  'use strict';

  const BRAND_NAME = 'VARGA CANTIERI';
  const BRAND_PATTERN = /\b(?:Hera App|Varga Cantieri)\b/gi;
  const SKIP_TAGS = new Set(['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEXTAREA', 'OPTION']);
  const BRAND_ATTRIBUTES = ['alt', 'aria-label', 'title'];

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

  function loadRubricaFeature() {
    try {
      if (document.querySelector('script[data-rubrica-feature]')) return;
      const script = document.createElement('script');
      script.src = './rubrica-feature.js?v=20260802-import1';
      script.defer = true;
      script.dataset.rubricaFeature = '1';
      script.addEventListener('error', () => console.warn('Rubrica non caricata; avvio app non interrotto.'), { once: true });
      document.head.appendChild(script);
    } catch (error) {
      console.warn('Rubrica non caricata; avvio app non interrotto:', error);
    }
  }

  function init() {
    applyBranding();
    observeDynamicContent();
    window.setTimeout(loadRubricaFeature, 0);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();

(() => {
  "use strict";

  if (window.HeraCantiereDocumentsLoaderInstalled) return;
  window.HeraCantiereDocumentsLoaderInstalled = true;

  const VERSION = "20260823-docs1";

  if (!document.querySelector('link[data-cantiere-doc-style]')) {
    const style = document.createElement("link");
    style.rel = "stylesheet";
    style.href = `documentazione-cantiere.css?v=${VERSION}`;
    style.dataset.cantiereDocStyle = "1";
    document.head.appendChild(style);
  }

  if (!document.querySelector('script[data-cantiere-doc-script]')) {
    const script = document.createElement("script");
    script.src = `documentazione-cantiere.js?v=${VERSION}`;
    script.async = false;
    script.dataset.cantiereDocScript = "1";
    script.onerror = () => console.error("Caricamento Documentazione cantiere non riuscito.");
    document.head.appendChild(script);
  }
})();
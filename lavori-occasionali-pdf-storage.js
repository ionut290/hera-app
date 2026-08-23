(() => {
  "use strict";

  if (window.HeraImpiantiPdfStorage?.installed) {
    window.HeraOccasionalPdfStorage = window.HeraImpiantiPdfStorage;
    return;
  }

  const src = "impianti-pdf-storage.js?v=20260823b";
  const existing = Array.from(document.scripts || []).find((script) => {
    try { return new URL(script.src, document.baseURI).pathname.endsWith("/impianti-pdf-storage.js"); }
    catch (_) { return false; }
  });

  const exposeAlias = () => {
    if (window.HeraImpiantiPdfStorage?.installed) {
      window.HeraOccasionalPdfStorage = window.HeraImpiantiPdfStorage;
    }
  };

  if (existing) {
    existing.addEventListener("load", exposeAlias, { once: true });
    exposeAlias();
    return;
  }

  const script = document.createElement("script");
  script.src = src;
  script.dataset.impiantiPdfStorage = "1";
  script.addEventListener("load", exposeAlias, { once: true });
  script.addEventListener("error", () => console.error("Modulo PDF impianti non caricato."), { once: true });
  document.head.appendChild(script);
})();

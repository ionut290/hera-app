/* Mantiene stabile la nuova vista Gestione impianti e contabilità.
   Se uno script non è disponibile dalla cache, lo ricarica prima di aprire la schermata. */
(() => {
  "use strict";

  let loadingPromise = null;

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = `${src}${src.includes("?") ? "&" : "?"}retry=${Date.now()}`;
      script.async = false;
      script.onload = resolve;
      script.onerror = () => reject(new Error(`Caricamento non riuscito: ${src}`));
      document.head.appendChild(script);
    });
  }

  async function ensureAccountingView() {
    if (window.InreteWorkItemsV2 && window.AccountingV2) return;
    if (loadingPromise) return loadingPromise;
    loadingPromise = (async () => {
      if (!window.InreteWorkItemsV2) {
        await loadScript("inrete-work-items-v2.js?v=20260728b");
      }
      if (!window.AccountingV2) {
        await loadScript("accounting-v2.js?v=20260901-current-gps1");
      }
      if (!window.InreteWorkItemsV2 || !window.AccountingV2) {
        throw new Error("La vista contabile non è disponibile.");
      }
    })().finally(() => {
      loadingPromise = null;
    });
    return loadingPromise;
  }

  window.openImpiantiManagement = async function openStableAccountingManagement(commessa) {
    try {
      await ensureAccountingView();
      return await window.AccountingV2.open(commessa);
    } catch (error) {
      console.error("Apertura Gestione impianti e contabilità non riuscita:", error);
      alert("Non è stato possibile caricare la contabilità. Controlla la connessione e riprova. La vecchia tabella non verrà aperta.");
      return null;
    }
  };
})();

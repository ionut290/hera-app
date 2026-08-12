(() => {
  "use strict";

  const originalSelectCommessa = window.selectCommessa;
  if (typeof originalSelectCommessa !== "function" || originalSelectCommessa.__navigationRepair) return;

  function forceCommessaNavigation(id) {
    try {
      localStorage.setItem("heraLastSelectedCommessaId", String(id || ""));
    } catch (_) {}

    try {
      window.stopImpiantiSubscription?.();
      window.stopCommessaNotesSubscription?.();

      const hasSubcommesse = typeof window.getSubcommesse === "function"
        && window.getSubcommesse(id).length > 0;

      if (!hasSubcommesse) {
        window.subscribeImpianti?.();
        window.subscribeCommessaNotes?.();
      }

      if (typeof window.setCommessaHash === "function") {
        window.setCommessaHash();
      } else {
        window.location.hash = `commessa=${encodeURIComponent(String(id || ""))}`;
      }
      window.applyRoute?.();
    } catch (fallbackError) {
      console.error("Ripristino apertura commessa non completato:", fallbackError);
      window.location.hash = `commessa=${encodeURIComponent(String(id || ""))}`;
    }
  }

  function selectCommessaWithNavigationRepair(id, nome, codice = "") {
    try {
      return originalSelectCommessa.call(this, id, nome, codice);
    } catch (error) {
      console.error("Errore durante apertura commessa; applico navigazione protetta:", {
        commessaId: id,
        commessaNome: nome,
        error
      });
      forceCommessaNavigation(id);
      return undefined;
    }
  }

  selectCommessaWithNavigationRepair.__navigationRepair = true;
  window.selectCommessa = selectCommessaWithNavigationRepair;
})();

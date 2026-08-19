/* Mantiene stabile la nuova vista Gestione impianti e contabilità.
   Se uno script non è disponibile dalla cache, lo ricarica prima di aprire la schermata.
   Ripara inoltre le commesse INRETE già marcate come migrate quando hanno impianti legacy ma 0 lavorazioni. */
(() => {
  "use strict";

  let loadingPromise = null;
  const recoveryInProgress = new Set();

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
        await loadScript("accounting-v2.js?v=20260812-modena2");
      }
      if (!window.InreteWorkItemsV2 || !window.AccountingV2) {
        throw new Error("La vista contabile non è disponibile.");
      }
    })().finally(() => {
      loadingPromise = null;
    });
    return loadingPromise;
  }

  async function repairEmptyInreteAccounting(commessa) {
    const id = String(commessa?.id || "");
    if (!id || recoveryInProgress.has(id)) return false;
    if (!window.InreteWorkItemsV2?.isInreteCommessa?.(commessa)) return false;
    if (typeof canManageData === "function" && !canManageData()) return false;
    if (!window.db || typeof getCommesseCollectionName !== "function") return false;

    const ref = db.collection(getCommesseCollectionName()).doc(id);
    const [workSnap, legacySnap] = await Promise.all([
      ref.collection("lavorazioni").limit(1).get(),
      ref.collection("impianti").limit(1).get()
    ]);
    if (!workSnap.empty || legacySnap.empty) return false;

    recoveryInProgress.add(id);
    try {
      console.warn(`[AccountingRecovery] ${commessa.nome || id}: impianti presenti ma lavorazioni assenti. Avvio ripristino migrazione INRETE v2.`);
      await ref.set({
        inreteMigrationVersion: 1,
        accountingRecoveryRequestedAt: firebase.firestore.FieldValue.serverTimestamp(),
        accountingRecoveryRequestedBy: currentUser?.uid || ""
      }, { merge: true });

      const repairedCommessa = {...commessa, inreteMigrationVersion: 1};
      await window.AccountingV2.open(repairedCommessa);

      const verifySnap = await ref.collection("lavorazioni").limit(1).get();
      if (verifySnap.empty) {
        throw new Error("Gli impianti esistono, ma la ricostruzione delle lavorazioni non ha prodotto righe.");
      }
      console.info(`[AccountingRecovery] ${commessa.nome || id}: ripristino completato.`);
      return true;
    } finally {
      recoveryInProgress.delete(id);
    }
  }

  window.openImpiantiManagement = async function openStableAccountingManagement(commessa) {
    try {
      await ensureAccountingView();
      await window.AccountingV2.open(commessa);
      try {
        await repairEmptyInreteAccounting(commessa);
      } catch (recoveryError) {
        console.error("Ripristino automatico contabilità INRETE non riuscito:", recoveryError);
        alert(`Gli impianti risultano presenti ma la contabilità non è stata ricostruita automaticamente. ${recoveryError?.message || recoveryError}`);
      }
      return true;
    } catch (error) {
      console.error("Apertura Gestione impianti e contabilità non riuscita:", error);
      alert("Non è stato possibile caricare la contabilità. Controlla la connessione e riprova. La vecchia tabella non verrà aperta.");
      return null;
    }
  };
})();

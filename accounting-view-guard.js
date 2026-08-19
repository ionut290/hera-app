/* Mantiene stabile la nuova vista Gestione impianti e contabilità.
   Se uno script non è disponibile dalla cache, lo ricarica prima di aprire la schermata.
   Ripara inoltre le commesse INRETE già marcate come migrate quando hanno impianti ma 0 lavorazioni. */
(() => {
  "use strict";

  let loadingPromise = null;
  const recoveryInProgress = new Set();
  const IMPIANTI_CACHE_PREFIX = "heraImpiantiPersistentCacheV1:";

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

  function cloneRows(rows) {
    return Array.isArray(rows) ? rows.map((row) => ({ ...(row || {}) })) : [];
  }

  function getMemoryPlants(commessaId) {
    const candidates = [];
    try {
      if (typeof getCommessaCachedImpianti === "function") {
        candidates.push(getCommessaCachedImpianti(commessaId));
      }
    } catch (_) {}
    try {
      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.get) {
        candidates.push(impiantiByCommessaId.get(commessaId));
      }
    } catch (_) {}
    try {
      const selectedId = String(typeof selectedCommessaId !== "undefined" ? selectedCommessaId : window.selectedCommessaId || "");
      if (selectedId === String(commessaId) && typeof currentImpianti !== "undefined") {
        candidates.push(currentImpianti);
      }
    } catch (_) {}
    return candidates.map(cloneRows).find((rows) => rows.length) || [];
  }

  function readPersistentPlants(commessaId) {
    const uid = String(currentUser?.uid || "").trim();
    if (!uid || !commessaId) return [];
    let collectionName = "commesse";
    try {
      collectionName = String(getCommesseCollectionName?.() || "commesse").trim() || "commesse";
    } catch (_) {}
    const key = `${IMPIANTI_CACHE_PREFIX}${encodeURIComponent(uid)}:${encodeURIComponent(collectionName)}:${encodeURIComponent(commessaId)}`;
    try {
      const parsed = JSON.parse(localStorage.getItem(key) || "null");
      if (!parsed || parsed.schemaVersion !== 1 || parsed.uid !== uid || parsed.collectionName !== collectionName
        || String(parsed.commessaId || "") !== String(commessaId) || !Array.isArray(parsed.items) || !parsed.items.length
        || !Number(parsed.markerMs)) return [];
      return cloneRows(parsed.items);
    } catch (_) {
      return [];
    }
  }

  function recoveryRows(commessaId) {
    const memory = getMemoryPlants(commessaId);
    if (memory.length) return { rows: memory, source: "memoria app" };
    const persistent = readPersistentPlants(commessaId);
    if (persistent.length) return { rows: persistent, source: "cache persistente verificata" };
    return { rows: [], source: "" };
  }

  async function restoreLegacyPlantsFromRecovery(ref, commessaId) {
    const existing = await ref.collection("impianti").limit(1).get();
    if (!existing.empty) return { restored: 0, source: "Firestore" };

    const recovery = recoveryRows(commessaId);
    if (!recovery.rows.length) return { restored: 0, source: "" };

    let restored = 0;
    for (let offset = 0; offset < recovery.rows.length; offset += 350) {
      const batch = db.batch();
      recovery.rows.slice(offset, offset + 350).forEach((row, index) => {
        const rawId = String(row.id || row.impiantoId || row.idSap || row.idSAP || "").trim();
        const fallbackId = `recovered_${offset + index + 1}`;
        const docId = rawId.replace(/[\/]/g, "_") || fallbackId;
        const data = { ...row };
        delete data.id;
        data.accountingRecoverySource = recovery.source;
        data.accountingRecoveryAt = firebase.firestore.FieldValue.serverTimestamp();
        batch.set(ref.collection("impianti").doc(docId), data, { merge: true });
        restored += 1;
      });
      await batch.commit();
    }
    console.warn(`[AccountingRecovery] ${commessaId}: ripristinati ${restored} impianti da ${recovery.source}.`);
    return { restored, source: recovery.source };
  }

  async function repairEmptyInreteAccounting(commessa) {
    const id = String(commessa?.id || "");
    if (!id || recoveryInProgress.has(id)) return false;
    if (!window.InreteWorkItemsV2?.isInreteCommessa?.(commessa)) return false;
    if (typeof canManageData === "function" && !canManageData()) return false;
    if (!window.db || typeof getCommesseCollectionName !== "function") return false;

    const ref = db.collection(getCommesseCollectionName()).doc(id);
    const workSnap = await ref.collection("lavorazioni").limit(1).get();
    if (!workSnap.empty) return false;

    recoveryInProgress.add(id);
    try {
      const restored = await restoreLegacyPlantsFromRecovery(ref, id);
      const legacySnap = await ref.collection("impianti").limit(1).get();
      if (legacySnap.empty) {
        console.warn(`[AccountingRecovery] ${commessa.nome || id}: nessun impianto disponibile in Firestore, memoria o cache persistente.`);
        return false;
      }

      console.warn(`[AccountingRecovery] ${commessa.nome || id}: lavorazioni assenti. Avvio ricostruzione INRETE v2.`);
      await ref.set({
        inreteMigrationVersion: 1,
        accountingRecoveryRequestedAt: firebase.firestore.FieldValue.serverTimestamp(),
        accountingRecoveryRequestedBy: currentUser?.uid || "",
        accountingRecoveryLegacyRestored: restored.restored || 0,
        accountingRecoverySource: restored.source || "Firestore"
      }, { merge: true });

      const repairedCommessa = { ...commessa, inreteMigrationVersion: 1 };
      await window.AccountingV2.open(repairedCommessa);

      const verifySnap = await ref.collection("lavorazioni").limit(1).get();
      if (verifySnap.empty) {
        throw new Error("La ricostruzione è partita ma non ha creato lavorazioni.");
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
        alert(`Ripristino automatico contabilità non completato: ${recoveryError?.message || recoveryError}`);
      }
      return true;
    } catch (error) {
      console.error("Apertura Gestione impianti e contabilità non riuscita:", error);
      alert("Non è stato possibile caricare la contabilità. Controlla la connessione e riprova. La vecchia tabella non verrà aperta.");
      return null;
    }
  };
})();

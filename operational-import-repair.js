/* Recupera manualmente gli impianti INRETE mancanti nella raccolta operativa.
   Nessuna scansione Firestore viene eseguita automaticamente all'avvio. */
(() => {
  "use strict";

  const REPAIR_VERSION = 2;
  const REPAIR_FIELD = "operationalRepairVersion";
  const inFlight = new Set();

  const parseNumber = (value) => {
    const parsed = Number.parseFloat(String(value ?? "").trim().replace(",", "."));
    return Number.isFinite(parsed) ? parsed : null;
  };
  const isDone = (item) => Boolean(item.done)
    || String(item.stato || "").trim().toUpperCase() === "FATTO";
  const isInrete = (commessa) => String(commessa?.nome || commessa?.codice || "")
    .trim().toLocaleUpperCase("it-IT").includes("INRETE");
  const collectionName = () => typeof getCommesseCollectionName === "function"
    ? getCommesseCollectionName()
    : "commesse";

  async function markChecked(ref, extra = {}) {
    await ref.set({
      [REPAIR_FIELD]: REPAIR_VERSION,
      operationalRepairCheckedAt: firebase.firestore.FieldValue.serverTimestamp(),
      ...extra
    }, { merge: true });
  }

  async function repairCommessa(commessa, { force = false } = {}) {
    const commessaId = String(commessa?.id || "").trim();
    if (!commessaId || inFlight.has(commessaId)) return false;
    if (!force && Number(commessa?.[REPAIR_FIELD] || 0) >= REPAIR_VERSION) return false;

    inFlight.add(commessaId);
    try {
      const commessaRef = db.collection(collectionName()).doc(commessaId);
      const operationalSnapshot = await commessaRef.collection("impianti").limit(1).get();
      if (!operationalSnapshot.empty) {
        await markChecked(commessaRef);
        return false;
      }

      const [workSnapshot, physicalSnapshot] = await Promise.all([
        commessaRef.collection("lavorazioni").get(),
        commessaRef.collection("impiantiFisici").get()
      ]);
      if (workSnapshot.empty || physicalSnapshot.empty) {
        await markChecked(commessaRef);
        return false;
      }

      const physicalById = new Map(
        physicalSnapshot.docs.map((doc) => [doc.id, { id: doc.id, ...doc.data() }])
      );
      const workByPlantId = new Map();
      workSnapshot.docs.forEach((doc) => {
        const work = { id: doc.id, ...doc.data() };
        const plantId = String(work.impiantoId || "").trim();
        if (!plantId || !physicalById.has(plantId)) return;
        const items = workByPlantId.get(plantId) || [];
        items.push(work);
        workByPlantId.set(plantId, items);
      });
      if (!workByPlantId.size) {
        await markChecked(commessaRef);
        return false;
      }

      const operations = [];
      let donePlants = 0;
      let doneWorkItems = 0;
      workByPlantId.forEach((items, plantId) => {
        const plant = physicalById.get(plantId);
        const completedItems = items.filter(isDone);
        const allDone = completedItems.length === items.length;
        const status = allDone ? "FATTO" : (completedItems.length ? "PARZIALMENTE FATTO" : "DA FARE");
        const lat = parseNumber(plant.latitudine ?? plant.gpsY);
        const lon = parseNumber(plant.longitudine ?? plant.gpsX);
        if (allDone) donePlants += 1;
        doneWorkItems += completedItems.length;
        operations.push({
          ref: commessaRef.collection("impianti").doc(plantId),
          data: {
            commessaId,
            physicalPlantId: plantId,
            migrationSourceId: plant.migrationSourceId || plantId,
            numeroProgressivo: plant.numeroProgressivoImpianto || plant.numeroProgressivo || null,
            distretto: plant.distretto || "",
            idSap: plant.idSap || "",
            denominazione: plant.denominazione || plant.nome || "",
            nome: plant.denominazione || plant.nome || "",
            comune: plant.comune || "",
            indirizzo: plant.indirizzo || plant.via || "",
            gpsY: lat,
            gpsX: lon,
            latitudine: lat,
            longitudine: lon,
            codicePrezzo: items.map((item) => String(item.codiceVocePrezzo || item.codicePrezzo || "").trim()).filter(Boolean).join("; "),
            tipologiaIntervento: items.map((item) => String(item.tipologiaLavorazione || item.tipologiaIntervento || "").trim()).filter(Boolean).join("; "),
            stato: status,
            statoGenerale: status,
            done: allDone,
            numeroLavorazioni: items.length,
            numeroLavorazioniFatte: completedItems.length,
            numeroLavorazioniDaFare: items.length - completedItems.length,
            repairedAt: firebase.firestore.FieldValue.serverTimestamp(),
            repairedBy: auth.currentUser?.uid || ""
          }
        });
      });

      for (let index = 0; index < operations.length; index += 400) {
        const batch = db.batch();
        operations.slice(index, index + 400).forEach(({ ref, data }) => batch.set(ref, data, { merge: true }));
        await batch.commit();
      }
      await markChecked(commessaRef, {
        impiantiCount: operations.length,
        totalPlants: operations.length,
        workItemsCount: workSnapshot.size,
        impiantiFattiCount: donePlants,
        impiantiDaFareCount: operations.length - donePlants,
        workItemsFattiCount: doneWorkItems,
        workItemsDaFareCount: workSnapshot.size - doneWorkItems,
        operationalModelSyncedAt: firebase.firestore.FieldValue.serverTimestamp()
      });
      console.info("[INRETE Import Repair] riparazione completata", { commessaId, impianti: operations.length, lavorazioni: workSnapshot.size });
      return true;
    } catch (error) {
      console.error("[INRETE Import Repair] riparazione non riuscita", { commessaId, error });
      return false;
    } finally {
      inFlight.delete(commessaId);
    }
  }

  async function runRepair({ force = false } = {}) {
    if (typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser) return false;
    if (typeof canManageData === "function" && !canManageData()) return false;
    const snapshot = await db.collection(collectionName()).get();
    const targets = snapshot.docs
      .map((doc) => ({ id: doc.id, ...doc.data() }))
      .filter(isInrete)
      .filter((commessa) => force || Number(commessa?.[REPAIR_FIELD] || 0) < REPAIR_VERSION);
    for (const commessa of targets) await repairCommessa(commessa, { force });
    return targets.length > 0;
  }

  function getSheetImportFeedback() {
    let feedback = document.querySelector("#sheet-url-feedback");
    if (feedback) return feedback;
    const row = document.querySelector("#sheet-url-import-btn")?.closest(".import-mode-details");
    if (!row) return document.querySelector("#import-feedback");
    feedback = document.createElement("p");
    feedback.id = "sheet-url-feedback";
    feedback.className = "muted";
    feedback.setAttribute("role", "status");
    feedback.setAttribute("aria-live", "polite");
    feedback.style.flexBasis = "100%";
    feedback.style.margin = "0";
    row.appendChild(feedback);
    return feedback;
  }

  function sheetImportEndpoint(sheetUrl) {
    const configuredOrigin = String(window.HERA_API_ORIGIN || "").trim();
    const localIsNetlify = /(?:^|\.)netlify\.app$/i.test(window.location.hostname);
    const apiOrigin = configuredOrigin || (localIsNetlify ? window.location.origin : "https://creative-syrniki-dddbae.netlify.app");
    const endpoint = new URL("/api/google-sheet-import", apiOrigin);
    endpoint.searchParams.set("url", sheetUrl);
    return endpoint.href;
  }

  async function readErrorMessage(response) {
    try {
      const payload = await response.json();
      return payload?.error || payload?.detail || `Errore HTTP ${response.status}`;
    } catch (_) {
      return `Errore HTTP ${response.status}`;
    }
  }

  async function importFromGoogleSheet(button) {
    const input = document.querySelector("#sheet-url");
    const sheetUrl = String(input?.value || "").trim();
    const feedback = getSheetImportFeedback();
    if (!sheetUrl) {
      if (feedback) feedback.textContent = "Incolla prima il link del Google Sheet.";
      input?.focus();
      return;
    }
    const originalText = button.textContent;
    button.disabled = true;
    button.textContent = "Caricamento…";
    if (feedback) feedback.textContent = "Lettura del Google Sheet in corso…";
    try {
      const response = await fetch(sheetImportEndpoint(sheetUrl), { method: "GET", headers: { Accept: "text/csv,application/json;q=0.9" }, cache: "no-store" });
      if (!response.ok) throw new Error(await readErrorMessage(response));
      const csv = await response.text();
      if (!csv.trim()) throw new Error("Il Google Sheet è vuoto.");
      const file = new File([csv], `google-sheet-${Date.now()}.csv`, { type: "text/csv;charset=utf-8", lastModified: Date.now() });
      const fileInput = document.querySelector("#excel-file");
      const importButton = document.querySelector("#import-btn");
      if (!fileInput || !importButton) throw new Error("Importazione matrice non disponibile in questa schermata.");
      const transfer = new DataTransfer();
      transfer.items.add(file);
      fileInput.files = transfer.files;
      fileInput.dispatchEvent(new Event("change", { bubbles: true }));
      if (feedback) feedback.textContent = "Foglio letto correttamente. Controlla l’anteprima e conferma l’importazione.";
      importButton.click();
    } catch (error) {
      const message = error?.message || "Errore durante la lettura del Google Sheet.";
      if (feedback) feedback.textContent = `${message} Verifica che il foglio sia condiviso come “Chiunque abbia il link – Visualizzatore”.`;
      console.error("[Google Sheet Import] importazione non riuscita", error);
    } finally {
      button.disabled = false;
      button.textContent = originalText;
    }
  }

  window.addEventListener("click", (event) => {
    const button = event.target?.closest?.("#sheet-url-import-btn");
    if (!button) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    void importFromGoogleSheet(button);
  }, true);

  // Disponibile solo come operazione manuale esplicita.
  window.repairImportedInretePlants = (options = {}) => runRepair({ force: options.force === true });
})();

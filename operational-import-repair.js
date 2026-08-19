/* Recupera manualmente gli impianti INRETE mancanti nella raccolta operativa.
   Include la sostituzione una-tantum degli impianti INRETE Modena con il dataset agosto 2026. */
(() => {
  "use strict";

  const REPAIR_VERSION = 2;
  const REPAIR_FIELD = "operationalRepairVersion";
  const MODENA_DATASET_VERSION = "2026-08-19-ago30-v1";
  const MODENA_DATASET_FIELD = "inreteModenaDatasetVersion";
  const inFlight = new Set();

  const MODENA_PLANTS = Object.freeze([
    { idSap: "3430707", denominazione: "REMI CASTELFRANCO", comune: "CASTELFRANCO EMILIA", indirizzo: "Via Loda, 28", lat: 44.58931, lng: 11.04352 },
    { idSap: "3430051", denominazione: "REMI S.CESARIO", comune: "S.CESARIO SUL PANARO", indirizzo: "Via della Cartiera, 33 (dietro al Cimitero)", lat: 44.56565, lng: 11.03273 },
    { idSap: "3472968", denominazione: "REMI SAVIGNANO", comune: "SAVIGNANO SUL PANARO", indirizzo: "Via S. Anna, 10", lat: 44.48468, lng: 11.03210 },
    { idSap: "3471231", denominazione: "REMI SPEZZANO", comune: "FIORANO MODENESE", indirizzo: "Via Crociale, 7", lat: 44.53632, lng: 10.83482 },
    { idSap: "3470478", denominazione: "REMI FORMIGINE", comune: "FORMIGINE", indirizzo: "Via Grandi", lat: 44.56882, lng: 10.84307 },
    { idSap: "3475570", denominazione: "REMI MAGRETA", comune: "SASSUOLO", indirizzo: "Via Secchia, 32", lat: 44.59869, lng: 10.79012 },
    { idSap: "3471127", denominazione: "REMI CASINALBO", comune: "FORMIGINE", indirizzo: "Via Sant Ambrogio", lat: 44.59549, lng: 10.84845 },
    { idSap: "3471152", denominazione: "REMI MARANELLO", comune: "MARANELLO", indirizzo: "Via Claudia", lat: 44.52638, lng: 10.85457 },
    { idSap: "3470982", denominazione: "REMI POZZA", comune: "MARANELLO", indirizzo: "Via Vandelli", lat: 44.53058, lng: 10.89367 },
    { idSap: "3471269", denominazione: "REMI BRAIDA", comune: "SASSUOLO", indirizzo: "Via S.Pietro/via San Bernardo", lat: 44.54703, lng: 10.79935 },
    { idSap: "area adiacente alla REMI", denominazione: "STOCCAGGIO TUBI BRAIDA", comune: "SASSUOLO", indirizzo: "Via S.Pietro/via San Bernardo", lat: 44.54703, lng: 10.79935 },
    { idSap: "3470559", denominazione: "REMI INDIPENDENZA", comune: "SASSUOLO", indirizzo: "Via Indipendenza", lat: 44.53781, lng: 10.77355 },
    { idSap: "3471102", denominazione: "REMI PONTE FOSSA", comune: "SASSUOLO", indirizzo: "Via Valle d'Aosta", lat: 44.56558, lng: 10.80002 },
    { idSap: "3471279", denominazione: "REMI UBERSETTO", comune: "FORMIGINE", indirizzo: "Via dei Prati", lat: 44.55092, lng: 10.86203 },
    { idSap: "3429985", denominazione: "REMI CASTELNUOVO", comune: "CASTELNUOVO RANGONE", indirizzo: "Via Gualinga, 23", lat: 44.53866, lng: 10.94932 },
    { idSap: "3430270", denominazione: "REMI CA' DI SOLA", comune: "CASTELVETRO", indirizzo: "Via Per Modena", lat: 44.53481, lng: 10.95387 },
    { idSap: "3430320", denominazione: "REMI SOLIGNANO", comune: "CASTELVETRO", indirizzo: "Via Montanara, 4", lat: 44.53274, lng: 10.91688 },
    { idSap: "3425784", denominazione: "REMI S.CLEMENTE", comune: "MODENA", indirizzo: "strada di S. Clemente, 11", lat: 44.70429, lng: 10.99211 },
    { idSap: "3426714", denominazione: "REMI S.CLEMENTE", comune: "MODENA", indirizzo: "strada di S. Clemente, 11", lat: 44.61267, lng: 10.97673 },
    { idSap: "3430248", denominazione: "REMI SUD", comune: "MODENA", indirizzo: "Via Cadiane, 255", lat: 44.60603, lng: 10.91114 },
    { idSap: "3430604", denominazione: "REMI CASTELLARO", comune: "SPILAMBERTO", indirizzo: "Via Castellaro, 13", lat: 44.54071, lng: 11.02894 },
    { idSap: "3430547", denominazione: "REMI S.VITO", comune: "SPILAMBERTO", indirizzo: "Via San Vito", lat: 44.53941, lng: 11.01387 },
    { idSap: "3430221", denominazione: "REMI VIGNOLA", comune: "VIGNOLA", indirizzo: "Via Doccia", lat: 44.48418, lng: 11.02692 },
    { idSap: "3476117", denominazione: "REMI BERZIGALA", comune: "SERRAMAZZONI", indirizzo: "Via Giardini (Loc. Ca Ambero)", lat: 44.39255, lng: 10.80896 },
    { idSap: "3473527", denominazione: "REMI PAVULLO CASTELLO", comune: "PAVULLO NEL FRIGNANO", indirizzo: "Via Montecuccolo", lat: 44.32698, lng: 10.83440 },
    { idSap: "3473528", denominazione: "REMI S.ANTONIO", comune: "PAVULLO NEL FRIGNANO", indirizzo: "Via Pico/Guicciardini", lat: 44.36767, lng: 10.83285 },
    { idSap: "3471537", denominazione: "REMI S. DALMAZIO", comune: "SERRAMAZZONI", indirizzo: "Via Per Marano", lat: 44.42214, lng: 10.85448 },
    { idSap: "3475075", denominazione: "REMI MONTECENERE", comune: "PAVULLO NEL FRIGNANO", indirizzo: "Via Bellini", lat: 44.31195, lng: 10.78122 },
    { idSap: "3426336", denominazione: "GRMI UMF07 _ EXPORT CERAM", comune: "MONTEFIORINO", indirizzo: "VIA LA PIANA, 6", lat: 44.37793, lng: 10.61998 },
    { idSap: "3426335", denominazione: "GRMI UMF08 _ EXPORT CERAM", comune: "MONTEFIORINO", indirizzo: "VIA LA PIANA, 6", lat: 44.37797, lng: 10.62001 }
  ]);

  const parseNumber = (value) => {
    const parsed = Number.parseFloat(String(value ?? "").trim().replace(",", "."));
    return Number.isFinite(parsed) ? parsed : null;
  };
  const isDone = (item) => Boolean(item.done) || String(item.stato || "").trim().toUpperCase() === "FATTO";
  const normalize = (value) => String(value ?? "").trim().toLocaleUpperCase("it-IT");
  const isInrete = (commessa) => normalize(commessa?.nome || commessa?.codice).includes("INRETE");
  const isInreteModena = (commessa) => {
    const haystack = [commessa?.nome, commessa?.name, commessa?.codice, commessa?.code, commessa?.cliente, commessa?.category, commessa?.categoria].map(normalize).join(" ");
    return haystack.includes("INRETE") && haystack.includes("MODENA");
  };
  const collectionName = () => typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";

  function plantDocId(plant, index) {
    const sap = String(plant.idSap || "").trim();
    if (/^\d+$/.test(sap)) return `sap_${sap}`;
    const slug = String(plant.denominazione || `impianto_${index + 1}`).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
    return `modena_${slug || index + 1}`;
  }

  async function deleteSnapshotDocs(snapshot) {
    for (let offset = 0; offset < snapshot.docs.length; offset += 400) {
      const batch = db.batch();
      snapshot.docs.slice(offset, offset + 400).forEach((doc) => batch.delete(doc.ref));
      await batch.commit();
    }
  }

  async function replaceInreteModenaDataset(commessa, { force = false } = {}) {
    const commessaId = String(commessa?.id || "").trim();
    if (!commessaId || !isInreteModena(commessa) || inFlight.has(`modena:${commessaId}`)) return false;
    if (!force && String(commessa?.[MODENA_DATASET_FIELD] || "") === MODENA_DATASET_VERSION) return false;
    inFlight.add(`modena:${commessaId}`);
    try {
      const commessaRef = db.collection(collectionName()).doc(commessaId);
      const [operationalSnapshot, physicalSnapshot] = await Promise.all([
        commessaRef.collection("impianti").get(),
        commessaRef.collection("impiantiFisici").get()
      ]);
      await deleteSnapshotDocs(operationalSnapshot);
      await deleteSnapshotDocs(physicalSnapshot);
      for (let offset = 0; offset < MODENA_PLANTS.length; offset += 200) {
        const batch = db.batch();
        MODENA_PLANTS.slice(offset, offset + 200).forEach((plant, localIndex) => {
          const index = offset + localIndex;
          const id = plantDocId(plant, index);
          const common = {
            commessaId, physicalPlantId: id, migrationSourceId: id,
            numeroProgressivo: index + 1, numeroProgressivoImpianto: index + 1,
            distretto: "Modena", idSap: plant.idSap,
            denominazione: plant.denominazione, nome: plant.denominazione,
            comune: plant.comune, indirizzo: plant.indirizzo, via: plant.indirizzo,
            gpsY: plant.lat, gpsX: plant.lng, latitudine: plant.lat, longitudine: plant.lng,
            stato: "DA FARE", statoGenerale: "DA FARE", done: false,
            sourceDataset: "06-Attività Sfalci - Ago-2026 - MO.xlsx",
            sourceDatasetVersion: MODENA_DATASET_VERSION,
            replacedAt: firebase.firestore.FieldValue.serverTimestamp(),
            replacedBy: auth.currentUser?.uid || ""
          };
          batch.set(commessaRef.collection("impianti").doc(id), common, { merge: false });
          batch.set(commessaRef.collection("impiantiFisici").doc(id), common, { merge: false });
        });
        await batch.commit();
      }
      await commessaRef.set({
        [MODENA_DATASET_FIELD]: MODENA_DATASET_VERSION,
        inreteModenaDatasetUpdatedAt: firebase.firestore.FieldValue.serverTimestamp(),
        impiantiCount: MODENA_PLANTS.length, totalPlants: MODENA_PLANTS.length,
        impiantiFattiCount: 0, impiantiDaFareCount: MODENA_PLANTS.length,
        operationalModelSyncedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
      console.info("[INRETE Modena] dataset sostituito", { commessaId, impianti: MODENA_PLANTS.length, version: MODENA_DATASET_VERSION });
      return true;
    } catch (error) {
      console.error("[INRETE Modena] sostituzione dataset non riuscita", { commessaId, error });
      return false;
    } finally {
      inFlight.delete(`modena:${commessaId}`);
    }
  }

  async function runModenaDatasetReplacement({ force = false } = {}) {
    if (typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser) return false;
    if (typeof canManageData === "function" && !canManageData()) return false;
    const snapshot = await db.collection(collectionName()).get();
    const targets = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() })).filter(isInreteModena);
    for (const commessa of targets) await replaceInreteModenaDataset(commessa, { force });
    return targets.length > 0;
  }

  async function markChecked(ref, extra = {}) {
    await ref.set({ [REPAIR_FIELD]: REPAIR_VERSION, operationalRepairCheckedAt: firebase.firestore.FieldValue.serverTimestamp(), ...extra }, { merge: true });
  }

  async function repairCommessa(commessa, { force = false } = {}) {
    const commessaId = String(commessa?.id || "").trim();
    if (!commessaId || inFlight.has(commessaId)) return false;
    if (!force && Number(commessa?.[REPAIR_FIELD] || 0) >= REPAIR_VERSION) return false;
    inFlight.add(commessaId);
    try {
      const commessaRef = db.collection(collectionName()).doc(commessaId);
      const operationalSnapshot = await commessaRef.collection("impianti").limit(1).get();
      if (!operationalSnapshot.empty) { await markChecked(commessaRef); return false; }
      const [workSnapshot, physicalSnapshot] = await Promise.all([commessaRef.collection("lavorazioni").get(), commessaRef.collection("impiantiFisici").get()]);
      if (workSnapshot.empty || physicalSnapshot.empty) { await markChecked(commessaRef); return false; }
      const physicalById = new Map(physicalSnapshot.docs.map((doc) => [doc.id, { id: doc.id, ...doc.data() }]));
      const workByPlantId = new Map();
      workSnapshot.docs.forEach((doc) => {
        const work = { id: doc.id, ...doc.data() };
        const plantId = String(work.impiantoId || "").trim();
        if (!plantId || !physicalById.has(plantId)) return;
        const items = workByPlantId.get(plantId) || []; items.push(work); workByPlantId.set(plantId, items);
      });
      if (!workByPlantId.size) { await markChecked(commessaRef); return false; }
      const operations = []; let donePlants = 0; let doneWorkItems = 0;
      workByPlantId.forEach((items, plantId) => {
        const plant = physicalById.get(plantId); const completedItems = items.filter(isDone);
        const allDone = completedItems.length === items.length;
        const status = allDone ? "FATTO" : (completedItems.length ? "PARZIALMENTE FATTO" : "DA FARE");
        const lat = parseNumber(plant.latitudine ?? plant.gpsY); const lon = parseNumber(plant.longitudine ?? plant.gpsX);
        if (allDone) donePlants += 1; doneWorkItems += completedItems.length;
        operations.push({ ref: commessaRef.collection("impianti").doc(plantId), data: {
          commessaId, physicalPlantId: plantId, migrationSourceId: plant.migrationSourceId || plantId,
          numeroProgressivo: plant.numeroProgressivoImpianto || plant.numeroProgressivo || null, distretto: plant.distretto || "",
          idSap: plant.idSap || "", denominazione: plant.denominazione || plant.nome || "", nome: plant.denominazione || plant.nome || "",
          comune: plant.comune || "", indirizzo: plant.indirizzo || plant.via || "", gpsY: lat, gpsX: lon, latitudine: lat, longitudine: lon,
          codicePrezzo: items.map((item) => String(item.codiceVocePrezzo || item.codicePrezzo || "").trim()).filter(Boolean).join("; "),
          tipologiaIntervento: items.map((item) => String(item.tipologiaLavorazione || item.tipologiaIntervento || "").trim()).filter(Boolean).join("; "),
          stato: status, statoGenerale: status, done: allDone, numeroLavorazioni: items.length,
          numeroLavorazioniFatte: completedItems.length, numeroLavorazioniDaFare: items.length - completedItems.length,
          repairedAt: firebase.firestore.FieldValue.serverTimestamp(), repairedBy: auth.currentUser?.uid || ""
        }});
      });
      for (let index = 0; index < operations.length; index += 400) {
        const batch = db.batch(); operations.slice(index, index + 400).forEach(({ ref, data }) => batch.set(ref, data, { merge: true })); await batch.commit();
      }
      await markChecked(commessaRef, { impiantiCount: operations.length, totalPlants: operations.length, workItemsCount: workSnapshot.size,
        impiantiFattiCount: donePlants, impiantiDaFareCount: operations.length - donePlants,
        workItemsFattiCount: doneWorkItems, workItemsDaFareCount: workSnapshot.size - doneWorkItems,
        operationalModelSyncedAt: firebase.firestore.FieldValue.serverTimestamp() });
      return true;
    } catch (error) {
      console.error("[INRETE Import Repair] riparazione non riuscita", { commessaId, error }); return false;
    } finally { inFlight.delete(commessaId); }
  }

  async function runRepair({ force = false } = {}) {
    if (typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser) return false;
    if (typeof canManageData === "function" && !canManageData()) return false;
    const snapshot = await db.collection(collectionName()).get();
    const targets = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() })).filter(isInrete).filter((commessa) => force || Number(commessa?.[REPAIR_FIELD] || 0) < REPAIR_VERSION);
    for (const commessa of targets) await repairCommessa(commessa, { force });
    return targets.length > 0;
  }

  function getSheetImportFeedback() {
    let feedback = document.querySelector("#sheet-url-feedback");
    if (feedback) return feedback;
    const row = document.querySelector("#sheet-url-import-btn")?.closest(".import-mode-details");
    if (!row) return document.querySelector("#import-feedback");
    feedback = document.createElement("p"); feedback.id = "sheet-url-feedback"; feedback.className = "muted";
    feedback.setAttribute("role", "status"); feedback.setAttribute("aria-live", "polite"); feedback.style.flexBasis = "100%"; feedback.style.margin = "0"; row.appendChild(feedback);
    return feedback;
  }

  function sheetImportEndpoint(sheetUrl) {
    const configuredOrigin = String(window.HERA_API_ORIGIN || "").trim();
    const localIsNetlify = /(?:^|\.)netlify\.app$/i.test(window.location.hostname);
    const apiOrigin = configuredOrigin || (localIsNetlify ? window.location.origin : "https://creative-syrniki-dddbae.netlify.app");
    const endpoint = new URL("/api/google-sheet-import", apiOrigin); endpoint.searchParams.set("url", sheetUrl); return endpoint.href;
  }

  async function readErrorMessage(response) {
    try { const payload = await response.json(); return payload?.error || payload?.detail || `Errore HTTP ${response.status}`; }
    catch (_) { return `Errore HTTP ${response.status}`; }
  }

  async function importFromGoogleSheet(button) {
    const input = document.querySelector("#sheet-url"); const sheetUrl = String(input?.value || "").trim(); const feedback = getSheetImportFeedback();
    if (!sheetUrl) { if (feedback) feedback.textContent = "Incolla prima il link del Google Sheet."; input?.focus(); return; }
    const originalText = button.textContent; button.disabled = true; button.textContent = "Caricamento…";
    if (feedback) feedback.textContent = "Lettura del Google Sheet in corso…";
    try {
      const response = await fetch(sheetImportEndpoint(sheetUrl), { method: "GET", headers: { Accept: "text/csv,application/json;q=0.9" }, cache: "no-store" });
      if (!response.ok) throw new Error(await readErrorMessage(response));
      const csv = await response.text(); if (!csv.trim()) throw new Error("Il Google Sheet è vuoto.");
      const file = new File([csv], `google-sheet-${Date.now()}.csv`, { type: "text/csv;charset=utf-8", lastModified: Date.now() });
      const fileInput = document.querySelector("#excel-file"); const importButton = document.querySelector("#import-btn");
      if (!fileInput || !importButton) throw new Error("Importazione matrice non disponibile in questa schermata.");
      const transfer = new DataTransfer(); transfer.items.add(file); fileInput.files = transfer.files; fileInput.dispatchEvent(new Event("change", { bubbles: true }));
      if (feedback) feedback.textContent = "Foglio letto correttamente. Controlla l’anteprima e conferma l’importazione."; importButton.click();
    } catch (error) {
      const message = error?.message || "Errore durante la lettura del Google Sheet.";
      if (feedback) feedback.textContent = `${message} Verifica che il foglio sia condiviso come “Chiunque abbia il link – Visualizzatore”.`;
      console.error("[Google Sheet Import] importazione non riuscita", error);
    } finally { button.disabled = false; button.textContent = originalText; }
  }

  window.addEventListener("click", (event) => {
    const button = event.target?.closest?.("#sheet-url-import-btn"); if (!button) return;
    event.preventDefault(); event.stopImmediatePropagation(); void importFromGoogleSheet(button);
  }, true);

  window.repairImportedInretePlants = (options = {}) => runRepair({ force: options.force === true });
  window.replaceInreteModenaAugust2026 = (options = {}) => runModenaDatasetReplacement({ force: options.force === true });

  let autoAttempts = 0;
  const tryAutoReplacement = async () => {
    autoAttempts += 1;
    try {
      if (typeof auth === "undefined" || typeof db === "undefined" || !auth.currentUser) {
        if (autoAttempts < 30) setTimeout(tryAutoReplacement, 1500); return;
      }
      if (typeof canManageData === "function" && !canManageData()) return;
      await runModenaDatasetReplacement();
    } catch (error) {
      console.error("[INRETE Modena] avvio automatico non riuscito", error);
      if (autoAttempts < 10) setTimeout(tryAutoReplacement, 3000);
    }
  };
  setTimeout(tryAutoReplacement, 1200);
})();

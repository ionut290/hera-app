/* Recupera automaticamente gli impianti INRETE importati nel modello contabile
   ma non ancora presenti nella raccolta operativa usata da mappa ed elenchi. */
(() => {
  "use strict";

  const inFlight = new Set();
  const parseNumber = (value) => {
    const parsed = Number.parseFloat(String(value ?? "").trim().replace(",", "."));
    return Number.isFinite(parsed) ? parsed : null;
  };
  const isDone = (item) => Boolean(item.done)
    || String(item.stato || "").trim().toUpperCase() === "FATTO";
  const isInrete = (commessa) => String(commessa?.nome || commessa?.codice || "")
    .trim()
    .toLocaleUpperCase("it-IT")
    .includes("INRETE");

  async function repairCommessa(commessa) {
    const commessaId = String(commessa?.id || "").trim();
    if (!commessaId || inFlight.has(commessaId)) return false;
    inFlight.add(commessaId);
    try {
      const collectionName = typeof getCommesseCollectionName === "function"
        ? getCommesseCollectionName()
        : "commesse";
      const commessaRef = db.collection(collectionName).doc(commessaId);
      const [workSnapshot, physicalSnapshot, operationalSnapshot] = await Promise.all([
        commessaRef.collection("lavorazioni").get(),
        commessaRef.collection("impiantiFisici").get(),
        commessaRef.collection("impianti").get()
      ]);
      if (!operationalSnapshot.empty || workSnapshot.empty || physicalSnapshot.empty) return false;

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
      if (!workByPlantId.size) return false;

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
            codicePrezzo: items
              .map((item) => String(item.codiceVocePrezzo || item.codicePrezzo || "").trim())
              .filter(Boolean)
              .join("; "),
            tipologiaIntervento: items
              .map((item) => String(item.tipologiaLavorazione || item.tipologiaIntervento || "").trim())
              .filter(Boolean)
              .join("; "),
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
        operations.slice(index, index + 400).forEach((operation) => {
          batch.set(operation.ref, operation.data, { merge: true });
        });
        await batch.commit();
      }
      await commessaRef.set({
        impiantiCount: operations.length,
        totalPlants: operations.length,
        workItemsCount: workSnapshot.size,
        impiantiFattiCount: donePlants,
        impiantiDaFareCount: operations.length - donePlants,
        workItemsFattiCount: doneWorkItems,
        workItemsDaFareCount: workSnapshot.size - doneWorkItems,
        operationalModelSyncedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
      console.info("[INRETE Import Repair] riparazione completata", {
        commessaId,
        impianti: operations.length,
        lavorazioni: workSnapshot.size
      });
      return true;
    } catch (error) {
      console.error("[INRETE Import Repair] riparazione non riuscita", { commessaId, error });
      return false;
    } finally {
      inFlight.delete(commessaId);
    }
  }

  async function runRepair() {
    if (typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser) return;
    if (typeof canManageData === "function" && !canManageData()) return;
    const collectionName = typeof getCommesseCollectionName === "function"
      ? getCommesseCollectionName()
      : "commesse";
    const snapshot = await db.collection(collectionName).get();
    const targets = snapshot.docs
      .map((doc) => ({ id: doc.id, ...doc.data() }))
      .filter(isInrete);
    for (const commessa of targets) await repairCommessa(commessa);
  }

  if (typeof auth !== "undefined") {
    auth.onAuthStateChanged((user) => {
      if (user) window.setTimeout(() => void runRepair(), 1200);
    });
  }
  window.addEventListener("online", () => void runRepair());
  window.repairImportedInretePlants = runRepair;
})();

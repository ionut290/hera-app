/* Crea INRETE MODENA - AGOSTO 2026 e la rende visibile anche se Firestore non completa la scrittura. */
(() => {
  "use strict";

  const COMMESSA_ID = "inrete_modena_agosto_2026";
  const DATASET_VERSION = "2026-08-19-ago30-ui-fallback-v2";
  const RETRY_MS = 2000;
  const MAX_ATTEMPTS = 90;
  let attempts = 0;
  let running = false;

  const RAW = Object.freeze([
    ["3430707","REMI CASTELFRANCO","CASTELFRANCO EMILIA","Via Loda, 28",44.58931,11.04352],
    ["3430051","REMI S.CESARIO","S.CESARIO SUL PANARO","Via della Cartiera, 33 (dietro al Cimitero)",44.56565,11.03273],
    ["3472968","REMI SAVIGNANO","SAVIGNANO SUL PANARO","Via S. Anna, 10",44.48468,11.03210],
    ["3471231","REMI SPEZZANO","FIORANO MODENESE","Via Crociale, 7",44.53632,10.83482],
    ["3470478","REMI FORMIGINE","FORMIGINE","Via Grandi",44.56882,10.84307],
    ["3475570","REMI MAGRETA","SASSUOLO","Via Secchia, 32",44.59869,10.79012],
    ["3471127","REMI CASINALBO","FORMIGINE","Via Sant Ambrogio",44.59549,10.84845],
    ["3471152","REMI MARANELLO","MARANELLO","Via Claudia",44.52638,10.85457],
    ["3470982","REMI POZZA","MARANELLO","Via Vandelli",44.53058,10.89367],
    ["3471269","REMI BRAIDA","SASSUOLO","Via S.Pietro/via San Bernardo",44.54703,10.79935],
    ["area adiacente alla REMI","STOCCAGGIO TUBI BRAIDA","SASSUOLO","Via S.Pietro/via San Bernardo",44.54703,10.79935],
    ["3470559","REMI INDIPENDENZA","SASSUOLO","Via Indipendenza",44.53781,10.77355],
    ["3471102","REMI PONTE FOSSA","SASSUOLO","Via Valle d'Aosta",44.56558,10.80002],
    ["3471279","REMI UBERSETTO","FORMIGINE","Via dei Prati",44.55092,10.86203],
    ["3429985","REMI CASTELNUOVO","CASTELNUOVO RANGONE","Via Gualinga, 23",44.53866,10.94932],
    ["3430270","REMI CA' DI SOLA","CASTELVETRO","Via Per Modena",44.53481,10.95387],
    ["3430320","REMI SOLIGNANO","CASTELVETRO","Via Montanara, 4",44.53274,10.91688],
    ["3425784","REMI S.CLEMENTE","MODENA","strada di S. Clemente, 11",44.70429,10.99211],
    ["3426714","REMI S.CLEMENTE","MODENA","strada di S. Clemente, 11",44.61267,10.97673],
    ["3430248","REMI SUD","MODENA","Via Cadiane, 255",44.60603,10.91114],
    ["3430604","REMI CASTELLARO","SPILAMBERTO","Via Castellaro, 13",44.54071,11.02894],
    ["3430547","REMI S.VITO","SPILAMBERTO","Via San Vito",44.53941,11.01387],
    ["3430221","REMI VIGNOLA","VIGNOLA","Via Doccia",44.48418,11.02692],
    ["3476117","REMI BERZIGALA","SERRAMAZZONI","Via Giardini (Loc. Ca Ambero)",44.39255,10.80896],
    ["3473527","REMI PAVULLO CASTELLO","PAVULLO NEL FRIGNANO","Via Montecuccolo",44.32698,10.83440],
    ["3473528","REMI S.ANTONIO","PAVULLO NEL FRIGNANO","Via Pico/Guicciardini",44.36767,10.83285],
    ["3471537","REMI S. DALMAZIO","SERRAMAZZONI","Via Per Marano",44.42214,10.85448],
    ["3475075","REMI MONTECENERE","PAVULLO NEL FRIGNANO","Via Bellini",44.31195,10.78122],
    ["3426336","GRMI UMF07 _ EXPORT CERAM","MONTEFIORINO","VIA LA PIANA, 6",44.37793,10.61998],
    ["3426335","GRMI UMF08 _ EXPORT CERAM","MONTEFIORINO","VIA LA PIANA, 6",44.37797,10.62001]
  ]);

  const plantId = (idSap, name, index) => /^\d+$/.test(String(idSap))
    ? `sap_${idSap}`
    : `modena_${String(name || index + 1).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "")}`;

  const PLANTS = RAW.map((row, index) => {
    const [idSap, denominazione, comune, indirizzo, lat, lng] = row;
    const id = plantId(idSap, denominazione, index);
    return {
      id,
      commessaId: COMMESSA_ID,
      physicalPlantId: id,
      migrationSourceId: id,
      numeroProgressivo: index + 1,
      numeroProgressivoImpianto: index + 1,
      distretto: "Modena",
      idSap,
      denominazione,
      nome: denominazione,
      comune,
      indirizzo,
      via: indirizzo,
      gpsY: lat,
      gpsX: lng,
      latitudine: lat,
      longitudine: lng,
      stato: "DA FARE",
      statoGenerale: "DA FARE",
      done: false,
      sourceDataset: "06-Attività Sfalci - Ago-2026 - MO.xlsx",
      sourceDatasetVersion: DATASET_VERSION
    };
  });

  const COMMESSA = {
    id: COMMESSA_ID,
    nome: "INRETE MODENA - AGOSTO 2026",
    name: "INRETE MODENA - AGOSTO 2026",
    codice: "INRETE-MO-AGO-2026",
    code: "INRETE-MO-AGO-2026",
    cliente: "INRETE",
    categoria: "INRETE GAS MODENA",
    area: "MODENA",
    stato: "Attiva",
    attiva: true,
    parentCommessaId: null,
    excelModelVersion: 2,
    priceListVersion: 2,
    workItemsModelVersion: 2,
    percentualeRibassoGenerale: 0.01,
    nextImpiantoNumber: 31,
    datasetVersion: DATASET_VERSION,
    sourceDataset: "06-Attività Sfalci - Ago-2026 - MO.xlsx",
    impiantiCount: PLANTS.length,
    totalPlants: PLANTS.length,
    impiantiFattiCount: 0,
    impiantiDaFareCount: PLANTS.length,
    localFallback: true
  };

  const collectionName = () => typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";

  function ensureVisibleLocally() {
    try {
      let changed = false;
      if (typeof commesseById !== "undefined" && commesseById?.set && !commesseById.has(COMMESSA_ID)) {
        commesseById.set(COMMESSA_ID, { ...COMMESSA });
        changed = true;
      }
      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.set) {
        const cached = impiantiByCommessaId.get(COMMESSA_ID);
        if (!Array.isArray(cached) || cached.length !== PLANTS.length) {
          impiantiByCommessaId.set(COMMESSA_ID, PLANTS.map((item) => ({ ...item })));
          changed = true;
        }
      }
      if (changed) {
        if (typeof renderCommesseHomeList === "function") renderCommesseHomeList();
        if (typeof renderCommesseManagementList === "function") renderCommesseManagementList();
        if (typeof refreshCommesseDependentUI === "function") refreshCommesseDependentUI(false);
      }
      if (typeof selectedCommessaId !== "undefined" && selectedCommessaId === COMMESSA_ID && typeof currentImpianti !== "undefined") {
        if (!Array.isArray(currentImpianti) || currentImpianti.length !== PLANTS.length) {
          currentImpianti = PLANTS.map((item) => ({ ...item }));
          if (typeof renderImpianti === "function") renderImpianti();
          if (typeof renderMap === "function") renderMap();
          if (typeof renderHeaderActivitySummary === "function") renderHeaderActivitySummary();
          if (typeof updateCommessaDashboard === "function") updateCommessaDashboard();
        }
      }
    } catch (error) {
      console.warn("[INRETE Modena fallback UI]", error);
    }
  }

  async function createInFirestore(force = false) {
    if (running || typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser) return false;
    running = true;
    try {
      const ref = db.collection(collectionName()).doc(COMMESSA_ID);
      const existing = await ref.get();
      const existingData = existing.exists ? (existing.data() || {}) : {};
      if (!force && existing.exists && String(existingData.datasetVersion || "") === DATASET_VERSION) return true;

      await ref.set({
        ...COMMESSA,
        localFallback: false,
        creatoDa: auth.currentUser.email || "",
        createdBy: auth.currentUser.uid,
        createdAt: existingData.createdAt || firebase.firestore.FieldValue.serverTimestamp(),
        updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
        operationalModelSyncedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });

      const batch = db.batch();
      PLANTS.forEach((plant) => {
        const data = {
          ...plant,
          createdAt: firebase.firestore.FieldValue.serverTimestamp(),
          updatedAt: firebase.firestore.FieldValue.serverTimestamp()
        };
        batch.set(ref.collection("impianti").doc(plant.id), data, { merge: true });
        batch.set(ref.collection("impiantiFisici").doc(plant.id), data, { merge: true });
      });
      await batch.commit();
      console.info("[INRETE Modena] nuova commessa salvata", { count: PLANTS.length });
      return true;
    } finally {
      running = false;
    }
  }

  async function tryRun() {
    ensureVisibleLocally();
    attempts += 1;
    try {
      if (typeof auth === "undefined" || typeof db === "undefined" || !auth.currentUser) throw new Error("Firebase non pronto");
      if (await createInFirestore(false)) return;
    } catch (error) {
      console.warn("[INRETE Modena] scrittura Firestore non riuscita, resta attivo il fallback UI", attempts, error?.message || error);
    }
    if (attempts < MAX_ATTEMPTS) setTimeout(tryRun, RETRY_MS);
  }

  window.createInreteModenaAugust2026 = (options = {}) => createInFirestore(options.force === true);
  window.INRETE_MODENA_AUGUST_2026 = { commessa: COMMESSA, plants: PLANTS, ensureVisibleLocally };

  setInterval(ensureVisibleLocally, 1500);
  if (typeof auth !== "undefined" && auth?.onAuthStateChanged) auth.onAuthStateChanged((user) => { if (user) setTimeout(tryRun, 250); });
  window.addEventListener("load", () => { ensureVisibleLocally(); setTimeout(tryRun, 500); });
  document.addEventListener("visibilitychange", () => { if (!document.hidden) ensureVisibleLocally(); });
  setTimeout(ensureVisibleLocally, 100);
  setTimeout(tryRun, 1000);
})();

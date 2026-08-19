/* INRETE operational repair + forced one-time Modena August 2026 replacement. */
(() => {
  "use strict";

  const MODENA_VERSION = "2026-08-19-ago30-v2";
  const MODENA_FIELD = "inreteModenaDatasetVersion";
  const RETRY_MS = 2000;
  const MAX_ATTEMPTS = 90;
  let attempts = 0;
  let running = false;

  const MODENA_PLANTS = Object.freeze([
    { idSap:"3430707", denominazione:"REMI CASTELFRANCO", comune:"CASTELFRANCO EMILIA", indirizzo:"Via Loda, 28", lat:44.58931, lng:11.04352 },
    { idSap:"3430051", denominazione:"REMI S.CESARIO", comune:"S.CESARIO SUL PANARO", indirizzo:"Via della Cartiera, 33 (dietro al Cimitero)", lat:44.56565, lng:11.03273 },
    { idSap:"3472968", denominazione:"REMI SAVIGNANO", comune:"SAVIGNANO SUL PANARO", indirizzo:"Via S. Anna, 10", lat:44.48468, lng:11.03210 },
    { idSap:"3471231", denominazione:"REMI SPEZZANO", comune:"FIORANO MODENESE", indirizzo:"Via Crociale, 7", lat:44.53632, lng:10.83482 },
    { idSap:"3470478", denominazione:"REMI FORMIGINE", comune:"FORMIGINE", indirizzo:"Via Grandi", lat:44.56882, lng:10.84307 },
    { idSap:"3475570", denominazione:"REMI MAGRETA", comune:"SASSUOLO", indirizzo:"Via Secchia, 32", lat:44.59869, lng:10.79012 },
    { idSap:"3471127", denominazione:"REMI CASINALBO", comune:"FORMIGINE", indirizzo:"Via Sant Ambrogio", lat:44.59549, lng:10.84845 },
    { idSap:"3471152", denominazione:"REMI MARANELLO", comune:"MARANELLO", indirizzo:"Via Claudia", lat:44.52638, lng:10.85457 },
    { idSap:"3470982", denominazione:"REMI POZZA", comune:"MARANELLO", indirizzo:"Via Vandelli", lat:44.53058, lng:10.89367 },
    { idSap:"3471269", denominazione:"REMI BRAIDA", comune:"SASSUOLO", indirizzo:"Via S.Pietro/via San Bernardo", lat:44.54703, lng:10.79935 },
    { idSap:"area adiacente alla REMI", denominazione:"STOCCAGGIO TUBI BRAIDA", comune:"SASSUOLO", indirizzo:"Via S.Pietro/via San Bernardo", lat:44.54703, lng:10.79935 },
    { idSap:"3470559", denominazione:"REMI INDIPENDENZA", comune:"SASSUOLO", indirizzo:"Via Indipendenza", lat:44.53781, lng:10.77355 },
    { idSap:"3471102", denominazione:"REMI PONTE FOSSA", comune:"SASSUOLO", indirizzo:"Via Valle d'Aosta", lat:44.56558, lng:10.80002 },
    { idSap:"3471279", denominazione:"REMI UBERSETTO", comune:"FORMIGINE", indirizzo:"Via dei Prati", lat:44.55092, lng:10.86203 },
    { idSap:"3429985", denominazione:"REMI CASTELNUOVO", comune:"CASTELNUOVO RANGONE", indirizzo:"Via Gualinga, 23", lat:44.53866, lng:10.94932 },
    { idSap:"3430270", denominazione:"REMI CA' DI SOLA", comune:"CASTELVETRO", indirizzo:"Via Per Modena", lat:44.53481, lng:10.95387 },
    { idSap:"3430320", denominazione:"REMI SOLIGNANO", comune:"CASTELVETRO", indirizzo:"Via Montanara, 4", lat:44.53274, lng:10.91688 },
    { idSap:"3425784", denominazione:"REMI S.CLEMENTE", comune:"MODENA", indirizzo:"strada di S. Clemente, 11", lat:44.70429, lng:10.99211 },
    { idSap:"3426714", denominazione:"REMI S.CLEMENTE", comune:"MODENA", indirizzo:"strada di S. Clemente, 11", lat:44.61267, lng:10.97673 },
    { idSap:"3430248", denominazione:"REMI SUD", comune:"MODENA", indirizzo:"Via Cadiane, 255", lat:44.60603, lng:10.91114 },
    { idSap:"3430604", denominazione:"REMI CASTELLARO", comune:"SPILAMBERTO", indirizzo:"Via Castellaro, 13", lat:44.54071, lng:11.02894 },
    { idSap:"3430547", denominazione:"REMI S.VITO", comune:"SPILAMBERTO", indirizzo:"Via San Vito", lat:44.53941, lng:11.01387 },
    { idSap:"3430221", denominazione:"REMI VIGNOLA", comune:"VIGNOLA", indirizzo:"Via Doccia", lat:44.48418, lng:11.02692 },
    { idSap:"3476117", denominazione:"REMI BERZIGALA", comune:"SERRAMAZZONI", indirizzo:"Via Giardini (Loc. Ca Ambero)", lat:44.39255, lng:10.80896 },
    { idSap:"3473527", denominazione:"REMI PAVULLO CASTELLO", comune:"PAVULLO NEL FRIGNANO", indirizzo:"Via Montecuccolo", lat:44.32698, lng:10.83440 },
    { idSap:"3473528", denominazione:"REMI S.ANTONIO", comune:"PAVULLO NEL FRIGNANO", indirizzo:"Via Pico/Guicciardini", lat:44.36767, lng:10.83285 },
    { idSap:"3471537", denominazione:"REMI S. DALMAZIO", comune:"SERRAMAZZONI", indirizzo:"Via Per Marano", lat:44.42214, lng:10.85448 },
    { idSap:"3475075", denominazione:"REMI MONTECENERE", comune:"PAVULLO NEL FRIGNANO", indirizzo:"Via Bellini", lat:44.31195, lng:10.78122 },
    { idSap:"3426336", denominazione:"GRMI UMF07 _ EXPORT CERAM", comune:"MONTEFIORINO", indirizzo:"VIA LA PIANA, 6", lat:44.37793, lng:10.61998 },
    { idSap:"3426335", denominazione:"GRMI UMF08 _ EXPORT CERAM", comune:"MONTEFIORINO", indirizzo:"VIA LA PIANA, 6", lat:44.37797, lng:10.62001 }
  ]);

  const normalize = (v) => String(v ?? "").trim().toLocaleUpperCase("it-IT");
  const collectionName = () => typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";
  const isModena = (c) => [c?.nome,c?.name,c?.codice,c?.code,c?.cliente,c?.category,c?.categoria,c?.descrizione]
    .map(normalize).join(" ").includes("INRETE") && [c?.nome,c?.name,c?.codice,c?.code,c?.cliente,c?.category,c?.categoria,c?.descrizione]
    .map(normalize).join(" ").includes("MODENA");

  function plantId(plant, index) {
    if (/^\d+$/.test(String(plant.idSap))) return `sap_${plant.idSap}`;
    return `modena_${String(plant.denominazione).normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^a-z0-9]+/g,"_").replace(/^_+|_+$/g,"") || index + 1}`;
  }

  async function deleteCollectionDocs(ref) {
    const snap = await ref.get();
    for (let i = 0; i < snap.docs.length; i += 400) {
      const batch = db.batch();
      snap.docs.slice(i, i + 400).forEach(doc => batch.delete(doc.ref));
      await batch.commit();
    }
  }

  async function replaceOne(commessa, force = false) {
    const id = String(commessa?.id || "").trim();
    if (!id || !isModena(commessa)) return false;
    if (!force && String(commessa?.[MODENA_FIELD] || "") === MODENA_VERSION) return true;

    const ref = db.collection(collectionName()).doc(id);
    await Promise.all([
      deleteCollectionDocs(ref.collection("impianti")),
      deleteCollectionDocs(ref.collection("impiantiFisici"))
    ]);

    for (let i = 0; i < MODENA_PLANTS.length; i += 200) {
      const batch = db.batch();
      MODENA_PLANTS.slice(i, i + 200).forEach((plant, j) => {
        const n = i + j;
        const docId = plantId(plant, n);
        const data = {
          commessaId:id,
          physicalPlantId:docId,
          migrationSourceId:docId,
          numeroProgressivo:n+1,
          numeroProgressivoImpianto:n+1,
          distretto:"Modena",
          idSap:plant.idSap,
          denominazione:plant.denominazione,
          nome:plant.denominazione,
          comune:plant.comune,
          indirizzo:plant.indirizzo,
          via:plant.indirizzo,
          gpsY:plant.lat,
          gpsX:plant.lng,
          latitudine:plant.lat,
          longitudine:plant.lng,
          stato:"DA FARE",
          statoGenerale:"DA FARE",
          done:false,
          sourceDataset:"06-Attività Sfalci - Ago-2026 - MO.xlsx",
          sourceDatasetVersion:MODENA_VERSION,
          replacedAt:firebase.firestore.FieldValue.serverTimestamp(),
          replacedBy:auth.currentUser?.uid || ""
        };
        batch.set(ref.collection("impianti").doc(docId), data);
        batch.set(ref.collection("impiantiFisici").doc(docId), data);
      });
      await batch.commit();
    }

    await ref.set({
      [MODENA_FIELD]:MODENA_VERSION,
      inreteModenaDatasetUpdatedAt:firebase.firestore.FieldValue.serverTimestamp(),
      impiantiCount:MODENA_PLANTS.length,
      totalPlants:MODENA_PLANTS.length,
      impiantiFattiCount:0,
      impiantiDaFareCount:MODENA_PLANTS.length,
      operationalModelSyncedAt:firebase.firestore.FieldValue.serverTimestamp()
    }, { merge:true });

    console.info("[INRETE Modena] sostituzione completata", { commessaId:id, impianti:MODENA_PLANTS.length, version:MODENA_VERSION });
    window.dispatchEvent(new CustomEvent("inrete-modena-dataset-replaced", { detail:{ commessaId:id, count:MODENA_PLANTS.length } }));
    return true;
  }

  async function run(force = false) {
    if (running || typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser) return false;
    running = true;
    try {
      const snap = await db.collection(collectionName()).get();
      const targets = snap.docs.map(d => ({ id:d.id, ...d.data() })).filter(isModena);
      if (!targets.length) throw new Error("Commessa INRETE Modena non trovata");
      for (const c of targets) await replaceOne(c, force);
      return true;
    } finally {
      running = false;
    }
  }

  window.replaceInreteModenaAugust2026 = (options = {}) => run(options.force === true);

  async function tryRun() {
    attempts += 1;
    try {
      if (typeof auth === "undefined" || typeof db === "undefined" || !auth.currentUser) throw new Error("Firebase non pronto");
      const ok = await run(false);
      if (ok) return;
    } catch (error) {
      console.warn("[INRETE Modena] tentativo migrazione", attempts, error?.message || error);
    }
    if (attempts < MAX_ATTEMPTS) setTimeout(tryRun, RETRY_MS);
  }

  if (typeof auth !== "undefined" && auth?.onAuthStateChanged) {
    auth.onAuthStateChanged(user => { if (user) setTimeout(tryRun, 250); });
  }
  window.addEventListener("load", () => setTimeout(tryRun, 500));
  document.addEventListener("visibilitychange", () => { if (!document.hidden) setTimeout(tryRun, 250); });
  setTimeout(tryRun, 1000);
})();

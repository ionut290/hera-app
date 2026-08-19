/* Inserisce i 30 impianti nella matrice della commessa INRETE Modena esistente. */
(() => {
  "use strict";

  const VERSION = "2026-08-19-inrete-modena-matrix-v1";
  const OLD_TEMP_COMMESSA_ID = "inrete_modena_agosto_2026";
  const RETRY_MS = 2000;
  const MAX_ATTEMPTS = 120;
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

  const normalize = (value) => String(value ?? "").trim().toLocaleUpperCase("it-IT");
  const collectionName = () => typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";

  function isRealInreteModena(commessa) {
    if (!commessa) return false;
    const id = String(commessa.id || "").trim();
    if (id === OLD_TEMP_COMMESSA_ID) return false;
    const text = normalize([
      commessa.nome, commessa.name, commessa.codice, commessa.code,
      commessa.cliente, commessa.categoria, commessa.category,
      commessa.descrizione, commessa.area
    ].filter(Boolean).join(" "));
    return text.includes("INRETE") && text.includes("MODENA") && !text.includes("AGOSTO 2026");
  }

  function plantDocId(idSap, name, index) {
    const sap = String(idSap || "").trim();
    if (/^\d+$/.test(sap)) return `sap_${sap}`;
    const slug = String(name || `impianto_${index + 1}`)
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
    return `modena_${slug || index + 1}`;
  }

  function buildPlants(commessaId) {
    return RAW.map((row, index) => {
      const [idSap, denominazione, comune, indirizzo, lat, lng] = row;
      const id = plantDocId(idSap, denominazione, index);
      return {
        id,
        commessaId,
        physicalPlantId: id,
        migrationSourceId: id,
        numeroProgressivo: index + 1,
        numeroProgressivoImpianto: index + 1,
        distretto: "Modena",
        areaCompetenza: "MODENA",
        idSap,
        "ID SAP": idSap,
        denominazione,
        "Denominazione Impianto": denominazione,
        nome: denominazione,
        comune,
        Comune: comune,
        indirizzo,
        via: indirizzo,
        "Descrizione via": indirizzo,
        gpsY: lat,
        gpsX: lng,
        "GPS(Y)": lat,
        "GPS(X)": lng,
        latitudine: lat,
        longitudine: lng,
        coordinate: `${lat}, ${lng}`,
        stato: "DA FARE",
        statoGenerale: "DA FARE",
        done: false,
        sourceDataset: "06-Attività Sfalci - Ago-2026 - MO.xlsx",
        sourceDatasetVersion: VERSION
      };
    });
  }

  function findLocalTarget() {
    if (typeof commesseById === "undefined" || !commesseById?.entries) return null;
    for (const [id, value] of commesseById.entries()) {
      const commessa = { id, ...(value || {}) };
      if (isRealInreteModena(commessa)) return commessa;
    }
    return null;
  }

  function applyMatrixLocally(target) {
    if (!target?.id) return false;
    const plants = buildPlants(target.id);
    try {
      if (typeof impiantiByCommessaId !== "undefined" && impiantiByCommessaId?.set) {
        const existing = Array.isArray(impiantiByCommessaId.get(target.id)) ? impiantiByCommessaId.get(target.id) : [];
        const map = new Map(existing.map((item) => [String(item.id || item.idSap || ""), item]));
        plants.forEach((plant) => map.set(String(plant.id), { ...map.get(String(plant.id)), ...plant }));
        impiantiByCommessaId.set(target.id, Array.from(map.values()));
      }
      if (typeof selectedCommessaId !== "undefined" && selectedCommessaId === target.id && typeof currentImpianti !== "undefined") {
        const existing = Array.isArray(currentImpianti) ? currentImpianti : [];
        const map = new Map(existing.map((item) => [String(item.id || item.idSap || ""), item]));
        plants.forEach((plant) => map.set(String(plant.id), { ...map.get(String(plant.id)), ...plant }));
        currentImpianti = Array.from(map.values());
        if (typeof renderImpianti === "function") renderImpianti();
        if (typeof renderMap === "function") renderMap();
        if (typeof renderImpiantiManagementTable === "function") renderImpiantiManagementTable();
        if (typeof updateCommessaDashboard === "function") updateCommessaDashboard();
      }
      return true;
    } catch (error) {
      console.warn("[INRETE Modena matrice] aggiornamento locale fallito", error);
      return false;
    }
  }

  async function findFirestoreTarget() {
    const snap = await db.collection(collectionName()).get();
    const targets = snap.docs.map((doc) => ({ id: doc.id, ...doc.data() })).filter(isRealInreteModena);
    if (!targets.length) return null;
    const exact = targets.find((item) => normalize(item.nome || item.name) === "INRETE MODENA")
      || targets.find((item) => normalize(item.codice || item.code).includes("MODENA"));
    return exact || targets[0];
  }

  async function persistMatrix(target, { force = false } = {}) {
    if (!target?.id || running) return false;
    running = true;
    try {
      const commessaRef = db.collection(collectionName()).doc(target.id);
      const current = await commessaRef.get();
      const currentData = current.exists ? (current.data() || {}) : {};
      if (!force && String(currentData.inreteModenaMatrixVersion || "") === VERSION) {
        applyMatrixLocally(target);
        return true;
      }

      const plants = buildPlants(target.id);
      const batch = db.batch();
      plants.forEach((plant) => {
        const payload = {
          ...plant,
          createdAt: firebase.firestore.FieldValue.serverTimestamp(),
          updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
          importedBy: auth.currentUser?.email || auth.currentUser?.uid || ""
        };
        batch.set(commessaRef.collection("impianti").doc(plant.id), payload, { merge: true });
        batch.set(commessaRef.collection("impiantiFisici").doc(plant.id), payload, { merge: true });
      });
      batch.set(commessaRef, {
        excelModelVersion: 2,
        nextImpiantoNumber: Math.max(Number(currentData.nextImpiantoNumber || 1), 31),
        inreteModenaMatrixVersion: VERSION,
        inreteModenaMatrixUpdatedAt: firebase.firestore.FieldValue.serverTimestamp(),
        impiantiCount: Math.max(Number(currentData.impiantiCount || 0), plants.length),
        totalPlants: Math.max(Number(currentData.totalPlants || 0), plants.length),
        updatedAt: firebase.firestore.FieldValue.serverTimestamp()
      }, { merge: true });
      await batch.commit();
      applyMatrixLocally(target);
      window.dispatchEvent(new CustomEvent("inrete-modena-matrix-updated", { detail: { commessaId: target.id, count: plants.length } }));
      console.info("[INRETE Modena matrice] 30 impianti inseriti", { commessaId: target.id, count: plants.length });
      return true;
    } finally {
      running = false;
    }
  }

  async function run(force = false) {
    const localTarget = findLocalTarget();
    if (localTarget) applyMatrixLocally(localTarget);
    if (typeof db === "undefined" || typeof auth === "undefined" || !auth.currentUser) return false;
    const target = localTarget || await findFirestoreTarget();
    if (!target) throw new Error("Commessa INRETE Modena esistente non trovata");
    return persistMatrix(target, { force });
  }

  window.insertInreteModenaPlantsIntoMatrix = (options = {}) => run(options.force === true);

  async function tryRun() {
    attempts += 1;
    try {
      if (await run(false)) return;
    } catch (error) {
      console.warn("[INRETE Modena matrice] tentativo", attempts, error?.message || error);
    }
    if (attempts < MAX_ATTEMPTS) setTimeout(tryRun, RETRY_MS);
  }

  if (typeof auth !== "undefined" && auth?.onAuthStateChanged) auth.onAuthStateChanged((user) => { if (user) setTimeout(tryRun, 250); });
  window.addEventListener("load", () => setTimeout(tryRun, 500));
  document.addEventListener("visibilitychange", () => { if (!document.hidden) setTimeout(tryRun, 250); });
  setTimeout(tryRun, 1000);
})();

#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const path = require("node:path");

const writes = [];
const reads = [];
const exportedFiles = [];
const timestamp = { kind: "SERVER_TIMESTAMP" };
const storageData = new Map();
const plant = {
  id: "albero-bologna-100",
  sourceIds: ["albero-bologna-100"],
  idSap: "ALB-BOLOGNA-100",
  denominazione: "Albero 100",
  gpsY: 44.4949,
  gpsX: 11.3426,
  potatureAbbattimenti: true,
  done: false
};

global.window = {
  addEventListener() {},
  alert() {},
  currentUser: { uid: "user-1", displayName: "Operatore Test" },
  localStorage: {
    getItem(key) { return storageData.get(key) || null; },
    setItem(key, value) { storageData.set(key, value); }
  }
};
Object.defineProperty(global, "navigator", { value: { onLine: true }, configurable: true });
global.document = {
  readyState: "loading",
  head: { appendChild() {} },
  addEventListener() {},
  getElementById() { return null; },
  querySelector() { return null; },
  querySelectorAll() { return []; },
  createElement() { return { id: "", textContent: "" }; }
};
global.selectedCommessaId = "potature-abbattimenti";
global.selectedCommessaName = "Potature Abbattimenti";
global.currentImpianti = [plant];
global.auth = { currentUser: { uid: "user-1", email: "operatore@example.test" } };
global.getCommesseCollectionName = () => "commesse";
global.getOperatorDisplayName = () => "Operatore Test";
global.buildRowsForEachCodicePrezzo = (item) => [{ ...item, cantiereRiga: item.denominazione }];
global.formatDoneDateTime = () => ({ date: "02/09/2026", time: "04:52" });
global.classifyTipoManutenzione = () => "Straordinaria";
global.XLSX = {
  utils: {
    json_to_sheet: (rows) => ({ rows }),
    book_new: () => ({ sheets: [] }),
    book_append_sheet: (workbook, worksheet, name) => workbook.sheets.push({ worksheet, name })
  },
  writeFile(workbook, filename) {
    exportedFiles.push({ workbook, filename });
  }
};
global.firebase = {
  firestore: {
    FieldValue: { serverTimestamp: () => timestamp },
    Timestamp: { fromDate: (date) => ({ kind: "CLIENT_TIMESTAMP", date }) }
  }
};
const activityEvents = [];
const notificationEvents = [];
global.logActivity = async (...args) => { activityEvents.push(args); };
global.publishGlobalNotificationEvent = async (...args) => { notificationEvents.push(args); };
global.db = {
  collection(collectionName) {
    assert.equal(collectionName, "commesse");
    return {
      doc(commessaId) {
        assert.equal(commessaId, "potature-abbattimenti");
        return {
          collection(subcollectionName) {
            assert.equal(subcollectionName, "impianti");
            return {
              doc: (id) => ({
                id,
                async get() {
                  reads.push(id);
                  const saved = [...writes].reverse().find((entry) => entry.reference.id === id)?.patch || {};
                  return { exists: true, data: () => saved };
                }
              })
            };
          }
        };
      }
    };
  },
  batch() {
    return {
      set(reference, patch, options) {
        writes.push({ reference, patch, options });
      },
      async commit() {}
    };
  }
};

const modulePath = path.resolve(__dirname, "..", "potature-followup.js");
delete require.cache[modulePath];
require(modulePath);

const api = global.window.HeraSpecialTerminato;
assert.equal(api?.installed, true, "API TERMINATO speciale installata");
assert.equal(api.isSpecialCommessa(), true, "commessa Potature riconosciuta come speciale");

(async () => {
  const button = { disabled: false, textContent: "TERMINATO" };
  const completionPromise = api.terminatePlant(plant, button);

  assert.equal(plant.specialTerminato, true, "passaggio locale nei FINITI immediato, prima della risposta Firestore");
  assert.equal(plant.specialTerminatoPending, true, "salvataggio marcato in attesa durante la transizione immediata");
  assert.equal(api.getDisplayState(plant).state, "Finito", "la vista considera subito il cantiere come finito");
  assert.equal(api.getDisplayState(plant).action, "FINITO", "TERMINATO non resta disponibile dopo il primo tocco");

  await completionPromise;

  assert.equal(writes.length, 1, "una sola scrittura per il documento del cantiere");
  assert.equal(writes[0].reference.id, "albero-bologna-100");
  assert.deepEqual(writes[0].options, { merge: true }, "scrittura non distruttiva");
  assert.equal(writes[0].patch.specialTerminato, true);
  assert.equal(writes[0].patch.specialTerminatoAt.kind, "CLIENT_TIMESTAMP");
  assert.equal(writes[0].patch.specialTerminatoBy, "Operatore Test");
  assert.equal(writes[0].patch.specialTerminatoByUid, "user-1");
  assert.equal(writes[0].patch.specialTerminatoByEmail, "operatore@example.test");
  assert.match(writes[0].patch.specialDataEsecuzione, /^\d{4}-\d{2}-\d{2}$/);
  assert.match(writes[0].patch.specialOraEsecuzione, /^\d{2}:\d{2}$/);
  assert.equal(writes[0].patch.specialStato, "FINITO");
  assert.equal(writes[0].patch.specialTerminatoPending, false);
  assert.equal(Object.hasOwn(writes[0].patch, "done"), false, "nessuna modifica allo stato FATTO");
  assert.equal(Object.hasOwn(writes[0].patch, "doneAt"), false, "nessuna modifica alla data FATTO");
  assert.equal(Object.hasOwn(writes[0].patch, "doneBy"), false, "nessuna modifica all’operatore FATTO");
  assert.equal(plant.specialTerminato, true, "stato locale speciale aggiornato");
  assert.equal(plant.done, false, "stato FATTO locale invariato");
  assert.equal(plant.specialTerminatoPending, false, "verifica online completata");
  assert.deepEqual(reads, ["albero-bologna-100"], "una verifica mirata dopo il batch");
  assert.equal(api.loadPendingActions().length, 0, "coda locale rimossa dopo la verifica");
  assert.equal(activityEvents.length, 1, "un solo evento attività TERMINATO");
  assert.equal(activityEvents[0][0], "pressione_terminato");
  assert.equal(notificationEvents.length, 1, "una sola notifica operativa TERMINATO");
  assert.equal(notificationEvents[0][0], "impianto-done");

  api.exportFinishedSummary();
  assert.equal(exportedFiles.length, 1, "un solo file Excel esportato");
  assert.match(exportedFiles[0].filename, /^riepilogo_impianti_Potature_Abbattimenti_/);
  assert.equal(exportedFiles[0].workbook.sheets[0].name, "Riepilogo impianti");
  const exportedRow = exportedFiles[0].workbook.sheets[0].worksheet.rows[0];
  assert.equal(exportedRow.Stato, "Finito");
  assert.equal(exportedRow["Data esecuzione"], "02/09/2026");
  assert.equal(exportedRow["Ora esecuzione"], "04:52");
  assert.equal(exportedRow["Eseguito da"], "Operatore Test");
  assert.equal(exportedRow["Email operatore"], "operatore@example.test");

  const offlinePlant = {
    id: "parco-cobo-200",
    sourceIds: ["parco-cobo-200"],
    idSap: "COBO-200",
    denominazione: "Parco COBO 200",
    gpsY: 44.505,
    gpsX: 11.31,
    done: false
  };
  global.currentImpianti.push(offlinePlant);
  global.navigator.onLine = false;
  const writesBeforeOffline = writes.length;
  await api.terminatePlant(offlinePlant, { disabled: false, textContent: "TERMINATO" });
  assert.equal(writes.length, writesBeforeOffline, "offline non avvia scritture Firestore inutili");
  assert.equal(offlinePlant.specialTerminato, true, "passaggio visivo immediato anche offline");
  assert.equal(offlinePlant.specialTerminatoPending, true, "stato offline marcato in attesa");
  assert.equal(offlinePlant.done, false, "FATTO resta invariato anche offline");
  assert.equal(api.loadPendingActions().length, 1, "azione offline conservata nella coda separata");

  global.navigator.onLine = true;
  await api.syncPendingActions();
  assert.equal(api.loadPendingActions().length, 0, "azione offline rimossa dopo la sincronizzazione");
  assert.equal(offlinePlant.specialTerminatoPending, false, "stato locale confermato dopo il ritorno online");
  assert.equal(writes.length, writesBeforeOffline + 1, "una sola scrittura al ritorno online");
  assert.deepEqual(reads, ["albero-bologna-100", "parco-cobo-200"], "una verifica mirata per ogni completamento");
  assert.equal(activityEvents.length, 2, "nessun evento attività duplicato durante la sincronizzazione");
  assert.equal(notificationEvents.length, 2, "nessuna notifica duplicata durante la sincronizzazione");

  const invalidPlant = { id: "senza-coordinate", sourceIds: ["senza-coordinate"], denominazione: "Senza coordinate", done: false };
  await assert.rejects(
    api.terminatePlant(invalidPlant, { disabled: false, textContent: "TERMINATO" }),
    /posizione.*mancante o non valida/i,
    "TERMINATO applica lo stesso controllo coordinate del flusso operativo"
  );
  assert.equal(invalidPlant.specialTerminato, undefined, "coordinate non valide non spostano il cantiere");

  console.log("✅ TERMINATO speciale: scrittura separata, esportazione Finiti e isolamento FATTO/WHAZZUP verificati.");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});

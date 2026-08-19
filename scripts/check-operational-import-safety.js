#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const vm = require("node:vm");
const assert = require("node:assert/strict");

const source = fs.readFileSync("operational-import-repair.js", "utf8");

function createRuntime(options = {}) {
  const writes = [];
  const deletes = [];
  const rootSets = [];
  const collectionCalls = [];
  const existingCommessa = options.existingCommessa ?? {
    exists: true,
    data: () => ({ datasetVersion: "old-version", nome: "Nome personalizzato", impiantiFattiCount: 9 })
  };
  const operationalDocs = options.operationalDocs || [];
  const physicalDocs = options.physicalDocs || [];
  const workDocs = options.workDocs || [];

  const makeDoc = (collectionName, id, data = {}) => ({
    id,
    data: () => ({ ...data }),
    ref: { collectionName, id }
  });
  const normalizeDocs = (collectionName, docs) => docs.map((entry) =>
    entry?.data && typeof entry.data === "function"
      ? entry
      : makeDoc(collectionName, entry.id, entry.data || entry)
  );

  const commessaRef = {
    get: async () => existingCommessa,
    set: async (payload, setOptions) => { rootSets.push({ payload, options: setOptions }); },
    collection(name) {
      const docs = name === "impianti"
        ? normalizeDocs(name, operationalDocs)
        : name === "impiantiFisici"
          ? normalizeDocs(name, physicalDocs)
          : normalizeDocs(name, workDocs);
      return {
        get: async () => ({ docs, size: docs.length, empty: docs.length === 0 }),
        doc(id) { return { collectionName: name, id }; }
      };
    }
  };

  const context = {
    console: { info() {}, warn() {}, error() {}, debug() {} },
    Map,
    Set,
    Promise,
    Date,
    Object,
    Array,
    String,
    Number,
    Boolean,
    RegExp,
    Math,
    JSON,
    setTimeout() { return 1; },
    clearTimeout() {},
    setInterval() { return 1; },
    clearInterval() {},
    document: {
      hidden: false,
      addEventListener() {}
    },
    firebase: {
      firestore: {
        FieldValue: { serverTimestamp: () => ({ __serverTimestamp: true }) }
      }
    },
    auth: {
      currentUser: { uid: "admin-1", email: "admin@example.test" },
      onAuthStateChanged() { return () => {}; }
    },
    canManageData: () => options.canManageData !== false,
    db: {
      collection(name) {
        collectionCalls.push(name);
        return { doc: () => commessaRef };
      },
      batch() {
        return {
          set(ref, payload, setOptions) { writes.push({ ref, payload, options: setOptions }); return this; },
          delete(ref) { deletes.push(ref); return this; },
          async commit() {}
        };
      }
    },
    commesseById: options.commesseById || new Map(),
    impiantiByCommessaId: options.impiantiByCommessaId || new Map(),
    selectedCommessaId: options.selectedCommessaId || "",
    currentImpianti: options.currentImpianti || [],
    renderImpianti() {},
    renderMap() {},
    renderHeaderActivitySummary() {},
    updateCommessaDashboard() {},
    renderCommesseHomeList() {},
    renderCommesseManagementList() {},
    refreshCommesseDependentUI() {},
    addEventListener() {}
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(source, context, { filename: "operational-import-repair.js" });
  return { context, api: context.INRETE_MODENA_AUGUST_2026, writes, deletes, rootSets, collectionCalls };
}

(async () => {
  {
    const { api } = createRuntime();
    const ordinary = api.detectKinds({ tipologiaLavorazione: "Manutenzione ordinaria" });
    const extraordinary = api.detectKinds({ tipologiaLavorazione: "Manutenzione straordinaria" });
    const mixed = api.detectKinds({ tipologiaLavorazione: "ORDINARIO E STRAORDINARIO" });
    assert.deepEqual([...ordinary], ["ORDINARIO"]);
    assert.deepEqual([...extraordinary], ["STRAORDINARIO"]);
    assert.equal(mixed.has("ORDINARIO"), true);
    assert.equal(mixed.has("STRAORDINARIO"), true);
  }

  {
    const { api } = createRuntime();
    const remoteDone = {
      ...api.plants[0],
      stato: "FATTO",
      statoGenerale: "FATTO",
      done: true,
      doneAt: "2026-08-19T08:00:00.000Z",
      operatore: "Cristina"
    };
    const merged = api.mergePlants(api.plants, [remoteDone]);
    const plant = merged.find((item) => item.id === remoteDone.id);
    assert.equal(plant.statoGenerale, "FATTO");
    assert.equal(plant.done, true);
    assert.equal(plant.doneAt, remoteDone.doneAt);
    assert.equal(plant.operatore, "Cristina");
  }

  {
    const runtime = createRuntime({
      workDocs: [{
        id: "work-extra",
        impiantoId: "sap_3430707",
        tipologiaLavorazione: "STRAORDINARIO",
        stato: "DA FARE"
      }]
    });
    await runtime.context.refreshInreteModenaMixedWork();
    const merged = runtime.api.mergePlants(runtime.api.plants);
    const plant = merged.find((item) => item.id === "sap_3430707");
    assert.equal(plant.attivitaLabel, "ORDINARIO E STRAORDINARIO");
    assert.equal(plant.isMixedOrdinaryExtraordinary, true);
    assert.equal(plant.markerClass, "impianto-marker-mixed-yellow");
  }

  {
    const doneData = {
      stato: "FATTO",
      statoGenerale: "FATTO",
      done: true,
      doneAt: "2026-08-18T10:00:00.000Z",
      tipo: "ORDINARIO E STRAORDINARIO",
      tipologiaIntervento: "ORDINARIO E STRAORDINARIO",
      attivitaLabel: "ORDINARIO E STRAORDINARIO",
      createdAt: "original-created-at"
    };
    const runtime = createRuntime({
      operationalDocs: [{ id: "sap_3430707", data: doneData }],
      physicalDocs: [{ id: "sap_3430707", data: doneData }]
    });
    const result = await runtime.context.createInreteModenaAugust2026({ force: true });
    assert.equal(result, true);
    assert.deepEqual([...new Set(runtime.collectionCalls)], ["commesse"]);

    const existingWrites = runtime.writes.filter((entry) => entry.ref.id === "sap_3430707");
    assert.equal(existingWrites.length, 2);
    for (const entry of existingWrites) {
      assert.equal(Object.hasOwn(entry.payload, "stato"), false);
      assert.equal(Object.hasOwn(entry.payload, "statoGenerale"), false);
      assert.equal(Object.hasOwn(entry.payload, "done"), false);
      assert.equal(Object.hasOwn(entry.payload, "tipo"), false);
      assert.equal(Object.hasOwn(entry.payload, "tipologiaIntervento"), false);
      assert.equal(Object.hasOwn(entry.payload, "attivitaLabel"), false);
      assert.equal(Object.hasOwn(entry.payload, "createdAt"), false);
    }

    assert.equal(runtime.rootSets.length, 1);
    const rootPayload = runtime.rootSets[0].payload;
    for (const field of ["nome", "codice", "parentCommessaId", "impiantiFattiCount", "impiantiDaFareCount", "percentualeRibassoGenerale", "nextImpiantoNumber"]) {
      assert.equal(Object.hasOwn(rootPayload, field), false, `campo commessa operativo sovrascritto: ${field}`);
    }
  }

  {
    const runtime = createRuntime({
      operationalDocs: [{ id: "sap_3430707", data: { denominazione: "REMI CASTELFRANCO" } }],
      physicalDocs: [{ id: "sap_3430707", data: { denominazione: "REMI CASTELFRANCO" } }]
    });
    await runtime.context.createInreteModenaAugust2026({ force: true });
    const writes = runtime.writes.filter((entry) => entry.ref.id === "sap_3430707");
    assert.equal(writes.every((entry) => entry.payload.stato === "DA FARE" && entry.payload.done === false), true);
  }

  {
    const runtime = createRuntime({ canManageData: false });
    const result = await runtime.context.createInreteModenaAugust2026({ force: true });
    assert.equal(result, false);
    assert.equal(runtime.collectionCalls.length, 0);
    assert.equal(runtime.writes.length, 0);
    assert.equal(runtime.rootSets.length, 0);
  }

  {
    const existing = new Map();
    const cachedDone = {
      id: "sap_3430707",
      idSap: "3430707",
      denominazione: "REMI CASTELFRANCO",
      stato: "FATTO",
      statoGenerale: "FATTO",
      done: true
    };
    const cached = new Map([["inrete_modena_agosto_2026", [cachedDone]]]);
    const runtime = createRuntime({
      commesseById: existing,
      impiantiByCommessaId: cached,
      selectedCommessaId: "inrete_modena_agosto_2026",
      currentImpianti: [cachedDone]
    });
    runtime.api.ensureVisibleLocally();
    assert.equal(runtime.context.currentImpianti.find((item) => item.id === cachedDone.id).done, true);
    assert.equal(runtime.context.impiantiByCommessaId.get("inrete_modena_agosto_2026").find((item) => item.id === cachedDone.id).done, true);
  }

  {
    const runtime = createRuntime();
    assert.equal(runtime.context.HeraGoogleSheetImportSafetyGuard?.installed, true);

    const batch = runtime.context.db.batch();
    batch.set(
      { path: "commesse/c1/lavorazioni/w1" },
      {
        sheetManaged: true,
        sheetSyncSourceId: "sheet:0",
        stato: "DA FARE",
        done: false,
        dataEsecuzione: "",
        oraEsecuzione: "",
        operatoreNome: "",
        note: "nota aggiornata"
      },
      { merge: true }
    );
    batch.delete({ path: "commesse/c1/lavorazioni/w2" });
    await batch.commit();

    const guardedWrite = runtime.writes.find((entry) => entry.ref.path === "commesse/c1/lavorazioni/w1");
    assert.equal(Object.hasOwn(guardedWrite.payload, "stato"), false);
    assert.equal(Object.hasOwn(guardedWrite.payload, "done"), false);
    assert.equal(Object.hasOwn(guardedWrite.payload, "dataEsecuzione"), false);
    assert.equal(guardedWrite.payload.note, "nota aggiornata");
    assert.equal(runtime.deletes.length, 0);

    const completedBatch = runtime.context.db.batch();
    completedBatch.set(
      { path: "commesse/c1/lavorazioni/w3" },
      {
        sheetManaged: true,
        sheetSyncSourceId: "sheet:0",
        stato: "FATTO",
        done: true,
        dataEsecuzione: "2026-08-19",
        oraEsecuzione: "10:30",
        operatoreNome: ""
      },
      { merge: true }
    );
    await completedBatch.commit();
    const completedWrite = runtime.writes.find((entry) => entry.ref.path === "commesse/c1/lavorazioni/w3");
    assert.equal(completedWrite.payload.stato, "FATTO");
    assert.equal(completedWrite.payload.done, true);
    assert.equal(completedWrite.payload.dataEsecuzione, "2026-08-19");
    assert.equal(Object.hasOwn(completedWrite.payload, "operatoreNome"), false);
  }

  {
    const runtime = createRuntime();
    await vm.runInContext(
      'db.batch().delete({ path: "globalCommesse/c2/impianti/sap_1" }).commit()',
      runtime.context,
      { filename: "google-sheet-two-way-sync.js" }
    );
    assert.equal(runtime.deletes.length, 0, "un foglio vuoto non deve cancellare impianti");
  }

  console.log("✅ Commesse/impianti INRETE Modena: stati FATTO preservati, lavorazioni collegate, scritture amministrative non distruttive.");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});

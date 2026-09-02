#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const path = require("node:path");

const writes = [];
const timestamp = { kind: "SERVER_TIMESTAMP" };
const plant = {
  id: "albero-bologna-100",
  sourceIds: ["albero-bologna-100"],
  idSap: "ALB-BOLOGNA-100",
  denominazione: "Albero 100",
  potatureAbbattimenti: true,
  done: false
};

global.window = {
  addEventListener() {},
  currentUser: { uid: "user-1", displayName: "Operatore Test" }
};
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
global.currentImpianti = [plant];
global.getCommesseCollectionName = () => "commesse";
global.getOperatorDisplayName = () => "Operatore Test";
global.firebase = {
  firestore: { FieldValue: { serverTimestamp: () => timestamp } }
};
global.db = {
  collection(collectionName) {
    assert.equal(collectionName, "commesse");
    return {
      doc(commessaId) {
        assert.equal(commessaId, "potature-abbattimenti");
        return {
          collection(subcollectionName) {
            assert.equal(subcollectionName, "impianti");
            return { doc: (id) => ({ id }) };
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
  await api.terminatePlant(plant, button);

  assert.equal(writes.length, 1, "una sola scrittura per il documento del cantiere");
  assert.equal(writes[0].reference.id, "albero-bologna-100");
  assert.deepEqual(writes[0].options, { merge: true }, "scrittura non distruttiva");
  assert.equal(writes[0].patch.specialTerminato, true);
  assert.equal(writes[0].patch.specialTerminatoAt, timestamp);
  assert.equal(writes[0].patch.specialTerminatoBy, "Operatore Test");
  assert.equal(writes[0].patch.specialTerminatoByUid, "user-1");
  assert.equal(Object.hasOwn(writes[0].patch, "done"), false, "nessuna modifica allo stato FATTO");
  assert.equal(Object.hasOwn(writes[0].patch, "doneAt"), false, "nessuna modifica alla data FATTO");
  assert.equal(Object.hasOwn(writes[0].patch, "doneBy"), false, "nessuna modifica all’operatore FATTO");
  assert.equal(plant.specialTerminato, true, "stato locale speciale aggiornato");
  assert.equal(plant.done, false, "stato FATTO locale invariato");

  console.log("✅ TERMINATO speciale: scrittura separata, non distruttiva e senza FATTO/WHAZZUP verificata.");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});

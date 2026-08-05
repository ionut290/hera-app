#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const path = require("node:path");

let physicalStarts = 0;
let physicalCloses = 0;
const nativeGroups = [];

function snapshotObserverFromArgs(args) {
  const list = Array.from(args);
  const first = list[0];
  if (first && typeof first === "object" && Object.prototype.hasOwnProperty.call(first, "includeMetadataChanges")) {
    return list[1];
  }
  if (typeof first === "function") {
    return { next: first, error: list[1], complete: list[2] };
  }
  return first;
}

class FakeQuery {
  constructor(queryPath, filtered = false) {
    if (!filtered) this.path = queryPath;
    this.nativePath = queryPath;
  }

  onSnapshot() {
    physicalStarts += 1;
    const observer = snapshotObserverFromArgs(arguments);
    const native = { observer, closed: false, path: this.nativePath };
    nativeGroups.push(native);
    return () => {
      if (native.closed) return;
      native.closed = true;
      physicalCloses += 1;
    };
  }
}

global.window = {
  firebase: { firestore: { Query: FakeQuery } },
  VargaFirestoreSafeOptimizer: { installed: true },
  addEventListener() {},
  setTimeout,
  clearTimeout
};

function emit(nativeIndex, value) {
  const native = nativeGroups[nativeIndex];
  assert.ok(native && !native.closed, `Listener fisico ${nativeIndex} non disponibile`);
  native.observer.next(value);
}

async function flushMicrotasks() {
  await Promise.resolve();
  await Promise.resolve();
}

async function run() {
  require(path.resolve(__dirname, "..", "firestore-nested-listener-optimizer.js"));
  const optimizer = window.VargaFirestoreNestedListenerOptimizer;
  assert.equal(optimizer.installed, true, "L'ottimizzatore impianti deve installarsi");

  const impiantiA = new FakeQuery("commesse/commessa-a/impianti");
  const impiantiB = new FakeQuery("commesse/commessa-a/impianti");
  const receivedA = [];
  const receivedB = [];

  const unsubscribeA = impiantiA.onSnapshot((snapshot) => receivedA.push(snapshot));
  const unsubscribeB = impiantiB.onSnapshot((snapshot) => receivedB.push(snapshot));
  assert.equal(physicalStarts, 1, "Due listener identici degli impianti devono usare un solo listener fisico");

  const firstSnapshot = { marker: "impianti-a-1" };
  emit(0, firstSnapshot);
  assert.deepEqual(receivedA, [firstSnapshot]);
  assert.deepEqual(receivedB, [firstSnapshot]);

  const otherCommessa = new FakeQuery("commesse/commessa-b/impianti");
  const unsubscribeOther = otherCommessa.onSnapshot(() => {});
  assert.equal(physicalStarts, 2, "Commesse diverse devono mantenere listener indipendenti");

  const filteredA = new FakeQuery("commesse/commessa-a/impianti", true);
  const filteredB = new FakeQuery("commesse/commessa-a/impianti", true);
  const unsubscribeFilteredA = filteredA.onSnapshot(() => {});
  const unsubscribeFilteredB = filteredB.onSnapshot(() => {});
  assert.equal(physicalStarts, 4, "Query filtrate devono mantenere il comportamento Firestore originale");

  unsubscribeA();
  unsubscribeB();

  const receivedC = [];
  const unsubscribeC = new FakeQuery("commesse/commessa-a/impianti")
    .onSnapshot((snapshot) => receivedC.push(snapshot));
  await flushMicrotasks();
  assert.equal(physicalStarts, 4, "La riapertura immediata deve riutilizzare il listener nel periodo di grazia");
  assert.deepEqual(receivedC, [firstSnapshot], "Il nuovo chiamante deve ricevere l'ultimo snapshot disponibile");

  unsubscribeC();
  unsubscribeOther();
  unsubscribeFilteredA();
  unsubscribeFilteredB();
  optimizer.clear();

  const state = optimizer.getState();
  assert.equal(state.activeGroups, 0);
  assert.equal(state.stats.preventedListenerStarts, 2, "Devono essere evitate la seconda apertura e la riapertura immediata");
  assert.equal(state.stats.gracePeriodReuses, 1);
  assert.equal(state.stats.physicalListenersStarted, 2);
  assert.equal(physicalCloses, 4, "Tutti i listener fisici devono poter essere chiusi normalmente");

  console.log("✅ I listener identici commesse/{id}/impianti condividono una sola apertura fisica.");
  console.log("✅ Commesse diverse e query filtrate restano indipendenti.");
  console.log("✅ La riapertura immediata riusa lo snapshot senza una nuova lettura iniziale.");
}

run().catch((error) => {
  console.error("❌ Test ottimizzatore listener impianti non superato.");
  console.error(error);
  process.exit(1);
});

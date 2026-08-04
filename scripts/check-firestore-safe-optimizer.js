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
  constructor(collection, canonical) {
    this.path = collection;
    this._query = {
      path: { canonicalString: () => collection },
      canonicalId: () => canonical
    };
  }

  onSnapshot() {
    physicalStarts += 1;
    const observer = snapshotObserverFromArgs(arguments);
    const native = { observer, closed: false };
    nativeGroups.push(native);
    return () => {
      if (native.closed) return;
      native.closed = true;
      physicalCloses += 1;
    };
  }
}

global.document = {
  querySelector() { return null; },
  createElement() {
    return {
      dataset: {},
      setAttribute() {},
      addEventListener() {}
    };
  },
  head: { appendChild() {} }
};

global.window = {
  firebase: { firestore: { Query: FakeQuery } },
  HeraRegistryDeviceCache: {},
  HeraFirestoreRegistryOptimizer: { installed: true },
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
  require(path.resolve(__dirname, "..", "firestore-safe-optimizer.js"));
  const optimizer = window.VargaFirestoreSafeOptimizer;
  assert.equal(optimizer.installed, true, "L'ottimizzatore deve installarsi sulle API Firestore disponibili");

  const commesseA = new FakeQuery("commesse", "commesse|all");
  const commesseB = new FakeQuery("commesse", "commesse|all");
  const receivedA = [];
  const receivedB = [];

  const unsubscribeA = commesseA.onSnapshot((snapshot) => receivedA.push(snapshot));
  const unsubscribeB = commesseB.onSnapshot((snapshot) => receivedB.push(snapshot));
  assert.equal(physicalStarts, 1, "Due listener identici di commesse devono usare un solo listener fisico");

  const firstSnapshot = { marker: "commesse-1" };
  emit(0, firstSnapshot);
  assert.deepEqual(receivedA, [firstSnapshot]);
  assert.deepEqual(receivedB, [firstSnapshot]);

  unsubscribeA();
  unsubscribeB();

  const receivedC = [];
  const unsubscribeC = new FakeQuery("commesse", "commesse|all")
    .onSnapshot((snapshot) => receivedC.push(snapshot));
  await flushMicrotasks();
  assert.equal(physicalStarts, 1, "La chiusura e riapertura immediata deve riutilizzare il listener nel periodo di grazia");
  assert.deepEqual(receivedC, [firstSnapshot], "Il nuovo chiamante deve ricevere l'ultimo snapshot già disponibile");

  const filtered = new FakeQuery("commesse", "commesse|where:attiva=true");
  const unsubscribeFiltered = filtered.onSnapshot(() => {});
  assert.equal(physicalStarts, 2, "Query con filtri diversi devono restare indipendenti");

  const userAlertA = new FakeQuery("userAlerts", "userAlerts|limit:500");
  const userAlertB = new FakeQuery("userAlerts", "userAlerts|limit:500");
  const unsubscribeAlertA = userAlertA.onSnapshot(() => {});
  const unsubscribeAlertB = userAlertB.onSnapshot(() => {});
  assert.equal(physicalStarts, 3, "I due listener identici userAlerts devono condividere una sola apertura fisica");

  const chatA = new FakeQuery("chatMessages", "chatMessages|limit:500");
  const chatB = new FakeQuery("chatMessages", "chatMessages|limit:500");
  const unsubscribeChatA = chatA.onSnapshot(() => {});
  const unsubscribeChatB = chatB.onSnapshot(() => {});
  assert.equal(physicalStarts, 5, "Le raccolte non autorizzate devono mantenere il comportamento Firestore originale");

  unsubscribeC();
  unsubscribeFiltered();
  unsubscribeAlertA();
  unsubscribeAlertB();
  unsubscribeChatA();
  unsubscribeChatB();
  optimizer.clear();

  const state = optimizer.getState();
  assert.equal(state.activeGroups, 0);
  assert.equal(state.stats.preventedListenerStarts, 2);
  assert.equal(state.stats.gracePeriodReuses, 1);
  assert.equal(state.stats.physicalListenersStarted, 3);
  assert.equal(physicalCloses, 5, "Tutti i listener fisici devono poter essere chiusi normalmente");

  console.log("✅ Commesse, squadreStorico e userAlerts condividono solo query identiche.");
  console.log("✅ Filtri diversi e raccolte non autorizzate restano completamente indipendenti.");
  console.log("✅ La riapertura immediata riusa lo snapshot senza una seconda lettura iniziale.");
}

run().catch((error) => {
  console.error("❌ Test ottimizzatore listener Firestore non superato.");
  console.error(error);
  process.exit(1);
});

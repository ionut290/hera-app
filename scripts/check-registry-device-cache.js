#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const path = require("node:path");

let networkGets = 0;
let profileWrites = 0;
let normalUpdates = 0;
let cacheMode = "valid";
let patchedCacheRecords = 0;

function networkSnapshot(query, records) {
  const docs = records.map((record) => ({
    id: record.id,
    ref: query.firestore.collection(query.path).doc(record.id),
    exists: true,
    data: () => ({ ...record, id: undefined })
  }));
  return {
    query,
    docs,
    size: docs.length,
    empty: docs.length === 0,
    metadata: { fromCache: false, hasPendingWrites: false },
    forEach(callback) { docs.forEach(callback); },
    docChanges() { return []; }
  };
}

class Query {
  constructor(collectionPath) {
    this.path = collectionPath;
    this.firestore = {
      collection: (name) => ({
        doc: (id) => ({ path: `${name}/${id}`, id })
      })
    };
  }

  get() {
    networkGets += 1;
    return Promise.resolve(networkSnapshot(this, [{ id: "server-1", name: "Dal server" }]));
  }

  onSnapshot(handler) {
    if (typeof handler === "function") handler(networkSnapshot(this, []));
    return () => {};
  }
}

class CollectionReference extends Query {
  add() {
    return Promise.resolve({ id: "new" });
  }
}

class DocumentReference {
  constructor(documentPath) {
    this.path = documentPath;
  }

  set() {
    profileWrites += 1;
    return Promise.resolve();
  }

  update() {
    normalUpdates += 1;
    return Promise.resolve();
  }

  delete() {
    return Promise.resolve();
  }
}

function firestore() {}
firestore.Query = Query;
firestore.CollectionReference = CollectionReference;
firestore.DocumentReference = DocumentReference;

global.window = {
  firebase: { firestore },
  HeraRegistryDeviceCache: {
    async readFresh(type) {
      if (cacheMode !== "valid") return null;
      if (type === "personale") {
        return { records: [{ id: "persona-1", name: "Mario", email: "mario@example.com" }], savedAt: Date.now() };
      }
      return { records: [{ id: "mezzo-1", targa: "AA000AA" }], savedAt: Date.now() };
    },
    async writeIfChanged() { return false; },
    async patchRecord() {
      patchedCacheRecords += 1;
      return true;
    },
    async remove() { return true; }
  }
};

global.CustomEvent = class CustomEvent {
  constructor(type, options) {
    this.type = type;
    this.detail = options?.detail;
  }
};

async function run() {
  require(path.resolve(__dirname, "..", "firestore-registry-read-optimizer.js"));

  const personale = new CollectionReference("personale");
  const cachedSnapshot = await personale.get();
  assert.equal(networkGets, 0, "Una cache valida non deve eseguire get() di rete");
  assert.equal(cachedSnapshot.size, 1);
  assert.equal(cachedSnapshot.docs[0].id, "persona-1");
  assert.equal(cachedSnapshot.docs[0].data().name, "Mario");
  assert.equal(cachedSnapshot.docs[0].ref.path, "personale/persona-1");
  assert.equal(cachedSnapshot.metadata.fromCache, true);

  const cachedAgain = await personale.get();
  assert.equal(networkGets, 0, "Il risultato recente deve essere riutilizzato");
  assert.equal(cachedAgain, cachedSnapshot);

  cacheMode = "missing";
  window.HeraFirestoreRegistryOptimizer.invalidate("mezzi");
  const mezzi = new CollectionReference("mezzi");
  const serverSnapshot = await mezzi.get();
  assert.equal(networkGets, 1, "Senza cache deve essere usata la query originale");
  assert.equal(serverSnapshot.metadata.fromCache, false);

  await mezzi.get({ source: "server" });
  assert.equal(networkGets, 2, "Le richieste con source esplicita non devono essere intercettate");

  const profileRef = new DocumentReference("personale/persona-1");
  const patch = { name: "Mario Rossi", email: "mario@example.com", updatedAt: { server: true } };
  await Promise.all([
    profileRef.set(patch, { merge: true }),
    profileRef.set(patch, { merge: true }),
    profileRef.set(patch, { merge: true }),
    profileRef.set(patch, { merge: true })
  ]);
  assert.equal(profileWrites, 1, "Patch profilo identiche devono condividere una sola scrittura");
  assert.equal(patchedCacheRecords, 1, "La scrittura profilo deve aggiornare anche IndexedDB");

  await profileRef.update({ ruolo: "operatore" });
  assert.equal(normalUpdates, 1);
  assert.equal(window.HeraFirestoreRegistryOptimizer.getState().stats.invalidations >= 2, true);

  const state = window.HeraFirestoreRegistryOptimizer.getState();
  assert.equal(state.stats.reusedDeviceCache >= 1, true);
  assert.equal(state.stats.profileWritesSkipped, 3);
  assert.equal(state.stats.profileWritesPassed, 1);
  assert.equal(state.stats.networkGets, 1);

  console.log("✅ Cache locale personale/mezzi: test automatici superati.");
  console.log("✅ Cache valida: nessun get di rete; fallback, invalidazione e profilo verificati.");
}

run().catch((error) => {
  console.error("❌ Test cache locale personale/mezzi non superato.");
  console.error(error);
  process.exit(1);
});

#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const path = require("node:path");

let networkGets = 0;
let listenerCalls = 0;
let cacheMode = "valid";
let cacheWrites = 0;

function makeFirestore() {
  return {
    collection(name) {
      return {
        doc(id) { return { path: `${name}/${id}`, id }; }
      };
    }
  };
}

function networkSnapshot(query, records) {
  const docs = records.map((record) => ({
    id: record.id,
    ref: query.firestore.collection(query.path).doc(record.id),
    exists: true,
    data: () => {
      const data = { ...record };
      delete data.id;
      return data;
    }
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
  constructor(collectionPath, shape = "query") {
    this.path = collectionPath;
    this.shape = shape;
    this.firestore = makeFirestore();
    this._query = { path: collectionPath };
  }

  orderBy(field, direction = "asc") {
    const query = new Query(this.path, `ordered:${field}:${direction}`);
    query.order = { field, direction };
    return query;
  }

  where() { return new Query(this.path, "filtered"); }
  limit() { return new Query(this.path, "limited"); }

  get() {
    networkGets += 1;
    const records = this.path === "personale"
      ? [
          { id: "server-2", name: "Server Due", createdAt: { seconds: 2 } },
          { id: "server-1", name: "Server Uno", createdAt: { seconds: 1 } }
        ]
      : [{ id: "mezzo-server", createdAt: { seconds: 1 } }];
    return Promise.resolve(networkSnapshot(this, records));
  }

  onSnapshot(handler) {
    listenerCalls += 1;
    const snapshot = networkSnapshot(this, [
      { id: "listener-1", name: "Listener", createdAt: { seconds: 1 } }
    ]);
    if (typeof handler === "function") handler(snapshot);
    else if (handler && typeof handler.next === "function") handler.next(snapshot);
    return () => {};
  }
}

class CollectionReference extends Query {
  constructor(collectionPath) {
    super(collectionPath, "collection");
  }

  doc(id) { return { path: `${this.path}/${id}`, id }; }
  add() { return Promise.resolve({ id: "new" }); }
}

function firestore() {}
firestore.Query = Query;
firestore.CollectionReference = CollectionReference;

global.window = {
  firebase: { firestore },
  addEventListener() {},
  HeraRegistryDeviceCache: {
    async readFresh(type) {
      if (cacheMode !== "valid") return null;
      if (type === "personale") {
        return {
          savedAt: Date.now(),
          records: [
            { id: "persona-2", name: "Seconda", createdAt: { seconds: 2 } },
            { id: "persona-1", name: "Prima", createdAt: { seconds: 1 } }
          ]
        };
      }
      return {
        savedAt: Date.now(),
        records: [{ id: "mezzo-1", createdAt: { seconds: 1 } }]
      };
    },
    async writeIfChanged() {
      cacheWrites += 1;
      return true;
    }
  }
};

async function run() {
  const nativeGet = Query.prototype.get;
  const existingGet = function existingRegistryGet() {
    return nativeGet.apply(this, arguments);
  };
  Object.defineProperty(existingGet, "__heraOriginal", { value: nativeGet });
  Query.prototype.get = existingGet;

  const nativeOnSnapshot = Query.prototype.onSnapshot;
  const existingOnSnapshot = function existingRegistryOnSnapshot() {
    return nativeOnSnapshot.apply(this, arguments);
  };
  Object.defineProperty(existingOnSnapshot, "__heraOriginal", { value: nativeOnSnapshot });
  Query.prototype.onSnapshot = existingOnSnapshot;

  window.HeraFirestoreRegistryOptimizer = {
    stats: {
      networkGets: 0,
      interceptedGets: 0,
      interceptedListeners: 0,
      reusedInFlight: 0,
      reusedRecent: 0,
      reusedDeviceCache: 0,
      sourceServerCacheHits: 0,
      networkFallbacks: 0,
      listenerSnapshots: 0,
      deviceCacheWrites: 0
    },
    getState() {
      return { installed: true, stats: { ...this.stats } };
    }
  };

  require(path.resolve(__dirname, "..", "firestore-registry-ordered-query-fix.js"));

  const personaleCollection = new CollectionReference("personale");
  const personaleQuery = personaleCollection.orderBy("createdAt", "asc");
  const cached = await personaleQuery.get({ source: "server" });

  assert.equal(networkGets, 0, "La query completa ordinata deve usare IndexedDB prima della rete");
  assert.equal(cached.metadata.fromCache, true);
  assert.deepEqual(cached.docs.map((doc) => doc.id), ["persona-1", "persona-2"], "La cache deve rispettare createdAt asc");

  const cachedAgain = await personaleQuery.get({ source: "server" });
  assert.equal(networkGets, 0, "La seconda lettura deve riutilizzare il risultato recente");
  assert.equal(cachedAgain, cached);

  const filtered = personaleCollection.orderBy("createdAt", "asc").where("ruolo", "==", "x");
  await filtered.get({ source: "server" });
  assert.equal(networkGets, 1, "Una query filtrata non deve ricevere l'intera cache personale");

  const limited = personaleCollection.orderBy("createdAt", "asc").limit(5);
  await limited.get({ source: "server" });
  assert.equal(networkGets, 2, "Una query con limit non deve ricevere l'intera cache personale");

  const otherOrder = personaleCollection.orderBy("nome", "asc");
  await otherOrder.get({ source: "server" });
  assert.equal(networkGets, 3, "Un ordinamento diverso non deve essere intercettato");

  cacheMode = "missing";
  window.HeraFirestoreRegistryOrderedQueryFix.invalidate("mezzi");
  const mezziQuery = new CollectionReference("mezzi").orderBy("createdAt", "asc");
  const server = await mezziQuery.get({ source: "server" });
  assert.equal(networkGets, 4, "Senza cache deve partire il fallback di rete");
  assert.equal(server.metadata.fromCache, false);

  let listenerSnapshot = null;
  personaleQuery.onSnapshot((snapshot) => { listenerSnapshot = snapshot; });
  await Promise.resolve();
  assert.equal(listenerCalls, 1, "Il listener originale deve restare attivo");
  assert.equal(listenerSnapshot.size, 1);
  assert.equal(cacheWrites >= 1, true, "Il listener deve aggiornare la cache locale");

  const optimizerState = window.HeraFirestoreRegistryOptimizer.getState();
  assert.equal(optimizerState.orderedQueryFix.installed, true);
  assert.equal(optimizerState.stats.interceptedGets >= 3, true);
  assert.equal(optimizerState.stats.reusedDeviceCache >= 1, true);
  assert.equal(optimizerState.stats.sourceServerCacheHits >= 2, true);
  assert.equal(optimizerState.orderedQueryFix.stats.markedQueries >= 4, true);

  console.log("✅ Query personale/mezzi ordinate per createdAt: cache verificata.");
  console.log("✅ where, limit e ordinamenti diversi: esclusi in sicurezza.");
  console.log("✅ Fallback rete, listener e diagnostica: verificati.");
}

run().catch((error) => {
  console.error("❌ Test query ordinate personale/mezzi non superato.");
  console.error(error);
  process.exit(1);
});

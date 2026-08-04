#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const optimizerSource = fs.readFileSync(
  path.resolve(__dirname, "..", "firestore-registry-read-optimizer.js"),
  "utf8"
);

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function createEnvironment({ firebaseReady = true } = {}) {
  const state = {
    networkGets: 0,
    lastGetOptions: undefined,
    profileWrites: 0,
    normalUpdates: 0,
    listenerCalls: 0,
    observerCalls: 0,
    cacheReads: 0,
    cacheWrites: 0,
    patchedCacheRecords: 0,
    cacheMode: "valid",
    cacheRecords: {
      personale: [{ id: "persona-1", name: "Mario", email: "mario@example.com" }],
      mezzi: [{ id: "mezzo-1", targa: "AA000AA" }]
    }
  };

  function networkSnapshot(query, records) {
    const collectionPath = query.path ||
      query?._query?.path?.canonicalString?.() ||
      query?.Ae?.path?.canonicalString?.() ||
      query?._delegate?._query?.path?.canonicalString?.() ||
      "personale";
    const docs = records.map((record) => ({
      id: record.id,
      ref: query.firestore.collection(collectionPath).doc(record.id),
      exists: true,
      metadata: { fromCache: false, hasPendingWrites: false },
      data: () => {
        const data = { ...record };
        delete data.id;
        return data;
      },
      get(fieldPath) {
        return String(fieldPath).split(".").reduce(
          (value, part) => value == null ? undefined : value[part],
          record
        );
      }
    }));
    return {
      query,
      docs,
      size: docs.length,
      empty: docs.length === 0,
      metadata: { fromCache: false, hasPendingWrites: false },
      forEach(callback, thisArg) {
        docs.forEach((doc) => callback.call(thisArg, doc));
      },
      docChanges() {
        return [];
      }
    };
  }

  class Query {
    constructor(collectionPath, internalVariant = "direct") {
      this.firestore = {
        collection: (name) => ({
          doc: (id) => ({ path: `${name}/${id}`, id })
        })
      };
      if (internalVariant === "direct") this.path = collectionPath;
      if (internalVariant === "_query") {
        this._query = { path: { canonicalString: () => collectionPath } };
      }
      if (internalVariant === "Ae") {
        this.Ae = { path: { canonicalString: () => collectionPath } };
      }
      if (internalVariant === "je") {
        this.je = { path: { segments: collectionPath.split("/") } };
      }
      if (internalVariant === "delegate") {
        this._delegate = {
          firestore: this.firestore,
          _query: { path: { canonicalString: () => collectionPath } }
        };
      }
    }

    get(options) {
      state.networkGets += 1;
      state.lastGetOptions = options;
      return Promise.resolve(networkSnapshot(this, [{ id: "server-1", name: "Dal server" }]));
    }

    onSnapshot() {
      const args = Array.from(arguments);
      const firstIsOptions = args[0] && typeof args[0] === "object" &&
        typeof args[0].next !== "function" &&
        Object.prototype.hasOwnProperty.call(args[0], "includeMetadataChanges");
      const handler = args[firstIsOptions ? 1 : 0];
      const snapshot = networkSnapshot(this, [{ id: "listener-1", name: "Listener" }]);
      if (typeof handler === "function") {
        state.listenerCalls += 1;
        handler(snapshot);
      } else if (handler && typeof handler.next === "function") {
        state.observerCalls += 1;
        handler.next(snapshot);
      }
      return () => {};
    }

    where() {
      return new Query(this.path || "personale", "direct");
    }
  }

  class CollectionReference extends Query {
    constructor(collectionPath, internalVariant = "direct") {
      super(collectionPath, internalVariant);
      this._collectionPathForTest = collectionPath;
    }

    doc(id) {
      return new DocumentReference(`${this._collectionPathForTest}/${id}`);
    }

    add() {
      return Promise.resolve({ id: "new" });
    }
  }

  class DocumentReference {
    constructor(documentPath) {
      this.path = documentPath;
    }

    set() {
      state.profileWrites += 1;
      return Promise.resolve();
    }

    update() {
      state.normalUpdates += 1;
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

  const listeners = new Map();
  const window = {
    firebase: firebaseReady ? { firestore } : undefined,
    HeraRegistryDeviceCache: {
      async readFresh(type) {
        state.cacheReads += 1;
        if (state.cacheMode === "throw") throw new Error("cache corrotta");
        if (state.cacheMode === "missing") return null;
        const records = state.cacheRecords[type];
        return records?.length ? { records, savedAt: Date.now() } : null;
      },
      async writeIfChanged() {
        state.cacheWrites += 1;
        return true;
      },
      async patchRecord() {
        state.patchedCacheRecords += 1;
        return true;
      },
      async remove() {
        return true;
      }
    },
    addEventListener(type, handler) {
      listeners.set(type, handler);
    },
    dispatchEvent(event) {
      const handler = listeners.get(event.type);
      if (handler) handler(event);
    },
    setTimeout,
    clearTimeout,
    console
  };

  const context = vm.createContext({
    window,
    console,
    setTimeout,
    clearTimeout,
    Promise,
    Date,
    Map,
    Set,
    WeakMap,
    Object,
    Array,
    String,
    Number,
    Boolean,
    JSON,
    Error
  });

  function installFirebase() {
    window.firebase = { firestore };
    const ready = listeners.get("hera:firebase-ready");
    if (ready) ready();
  }

  return {
    state,
    window,
    context,
    classes: { Query, CollectionReference, DocumentReference },
    installFirebase
  };
}

async function runOptimizer(env) {
  vm.runInContext(optimizerSource, env.context, {
    filename: "firestore-registry-read-optimizer.js"
  });
  await delay(0);
}

async function testDirectAndSourceVariants() {
  const env = createEnvironment();
  await runOptimizer(env);
  const { CollectionReference } = env.classes;

  const personale = new CollectionReference("personale");
  const cached = await personale.get({ source: "server" });
  assert.equal(env.state.networkGets, 0, "source=server deve usare la cache valida prima della rete");
  assert.equal(cached.docs[0].id, "persona-1");
  assert.equal(cached.metadata.fromCache, true);
  assert.equal(cached.docs[0].ref.path, "personale/persona-1");
  assert.equal(cached.docs[0].get("name"), "Mario");
  assert.equal(cached.docs[0].data().id, undefined, "l'ID tecnico non deve entrare nei dati Firestore");

  const repeated = await personale.get({ source: "server" });
  assert.equal(env.state.networkGets, 0, "la lettura recente deve essere riutilizzata");
  assert.equal(repeated, cached);

  await personale.get({ source: "cache" });
  assert.equal(env.state.networkGets, 1, "source=cache deve restare sotto il controllo nativo di Firestore");
  assert.deepEqual(env.state.lastGetOptions, { source: "cache" });

  await personale.get({ source: "future-option", custom: true });
  assert.equal(env.state.networkGets, 2, "opzioni sconosciute devono usare il comportamento nativo");

  const stats = env.window.HeraFirestoreRegistryOptimizer.getState().stats;
  assert.equal(stats.interceptedGets, 2);
  assert.equal(stats.sourceServerCacheHits >= 2, true);
  assert.equal(stats.reusedDeviceCache >= 1, true);
  assert.equal(stats.reusedRecent >= 1, true);
  assert.equal(stats.sourceCacheBypasses, 1);
  assert.equal(stats.unsupportedOptionBypasses, 1);
}

async function testNetworkFallbackPreservesOptionsAndErrors() {
  const env = createEnvironment();
  env.state.cacheMode = "missing";
  await runOptimizer(env);
  const mezzi = new env.classes.CollectionReference("mezzi");

  const snapshot = await mezzi.get({ source: "server" });
  assert.equal(env.state.networkGets, 1);
  assert.deepEqual(env.state.lastGetOptions, { source: "server" });
  assert.equal(snapshot.metadata.fromCache, false);

  env.window.HeraFirestoreRegistryOptimizer.invalidate("mezzi");
  env.state.cacheMode = "throw";
  await mezzi.get({ source: "server" });
  assert.equal(env.state.networkGets, 2, "errore IndexedDB deve fare fallback alla rete");
  const stats = env.window.HeraFirestoreRegistryOptimizer.getState().stats;
  assert.equal(stats.deviceCacheReadErrors, 1);
  assert.equal(stats.networkGets, 2);
  assert.equal(stats.networkFallbacks, 2);
}

async function testFirebaseV8InternalVariants() {
  for (const variant of ["_query", "Ae", "je", "delegate"]) {
    const env = createEnvironment();
    await runOptimizer(env);
    const ref = new env.classes.CollectionReference("personale", variant);
    const snapshot = await ref.get({ source: "server" });
    assert.equal(
      env.state.networkGets,
      0,
      `la variante interna ${variant} deve essere riconosciuta`
    );
    assert.equal(snapshot.docs[0].id, "persona-1");
    assert.equal(
      env.window.HeraFirestoreRegistryOptimizer.getState().stats.interceptedGets,
      1
    );
  }
}

async function testFilteredQueriesAreNeverServedAsFullCollections() {
  const env = createEnvironment();
  await runOptimizer(env);
  const filtered = new env.classes.Query("personale", "_query");
  const snapshot = await filtered.get({ source: "server" });
  assert.equal(env.state.networkGets, 1, "una query filtrata non deve ricevere l'intera cache personale");
  assert.equal(snapshot.docs[0].id, "server-1");
  assert.equal(
    env.window.HeraFirestoreRegistryOptimizer.getState().stats.filteredQueryBypasses,
    1
  );
}

async function testListenersCallbacksAndObservers() {
  const env = createEnvironment();
  await runOptimizer(env);

  const personale = new env.classes.CollectionReference("personale", "Ae");
  let callbackSnapshot = null;
  personale.onSnapshot({ includeMetadataChanges: true }, (snapshot) => {
    callbackSnapshot = snapshot;
  });
  assert.equal(callbackSnapshot.docs[0].id, "listener-1");
  assert.equal(env.state.listenerCalls, 1);

  const mezzi = new env.classes.CollectionReference("mezzi", "delegate");
  let observerSnapshot = null;
  mezzi.onSnapshot({
    next(snapshot) {
      observerSnapshot = snapshot;
    },
    error() {}
  });
  assert.equal(observerSnapshot.docs[0].id, "listener-1");
  assert.equal(env.state.observerCalls, 1);

  await delay(0);
  const stats = env.window.HeraFirestoreRegistryOptimizer.getState().stats;
  assert.equal(stats.interceptedListeners, 2);
  assert.equal(stats.listenerSnapshots, 2);
  assert.equal(env.state.cacheWrites, 2);
}

async function testLateFirebaseInstallation() {
  const env = createEnvironment({ firebaseReady: false });
  await runOptimizer(env);
  assert.equal(env.window.HeraFirestoreRegistryOptimizer.installed, false);

  env.installFirebase();
  await delay(0);
  assert.equal(env.window.HeraFirestoreRegistryOptimizer.installed, true);

  const personale = new env.classes.CollectionReference("personale", "_query");
  await personale.get({ source: "server" });
  assert.equal(env.state.networkGets, 0);
  assert.equal(
    env.window.HeraFirestoreRegistryOptimizer.getState().stats.patchedPrototypeCount >= 1,
    true
  );
}

async function testMutationSafetyAndDedupe() {
  const env = createEnvironment();
  await runOptimizer(env);

  const profileRef = new env.classes.DocumentReference("personale/persona-1");
  const patch = {
    name: "Mario Rossi",
    email: "mario@example.com",
    updatedAt: { server: true }
  };
  await Promise.all([
    profileRef.set(patch, { merge: true }),
    profileRef.set(patch, { merge: true }),
    profileRef.set(patch, { merge: true }),
    profileRef.set(patch, { merge: true })
  ]);
  assert.equal(env.state.profileWrites, 1, "le patch profilo identiche devono condividere una sola scrittura");
  assert.equal(env.state.patchedCacheRecords, 1);

  await profileRef.update({ ruolo: "operatore" });
  assert.equal(env.state.normalUpdates, 1, "gli aggiornamenti normali non devono essere soppressi");

  const stats = env.window.HeraFirestoreRegistryOptimizer.getState().stats;
  assert.equal(stats.profileWritesSkipped, 3);
  assert.equal(stats.profileWritesPassed, 1);
  assert.equal(stats.invalidations >= 1, true);
}

async function run() {
  await testDirectAndSourceVariants();
  await testNetworkFallbackPreservesOptionsAndErrors();
  await testFirebaseV8InternalVariants();
  await testFilteredQueriesAreNeverServedAsFullCollections();
  await testListenersCallbacksAndObservers();
  await testLateFirebaseInstallation();
  await testMutationSafetyAndDedupe();

  console.log("✅ Ottimizzatore personale/mezzi: tutte le varianti automatiche superate.");
  console.log("✅ Verificati Firebase v8, source server/cache, fallback rete, listener, query filtrate e caricamento tardivo.");
  console.log("✅ Verificata la protezione delle scritture profilo senza modificare dati, ore, calendario, squadre, FATTO o WhatsApp.");
}

run().catch((error) => {
  console.error("❌ Test ottimizzatore personale/mezzi non superato.");
  console.error(error);
  process.exit(1);
});
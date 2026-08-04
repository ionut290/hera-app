#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const path = require("node:path");

let calls = 0;
let rejectNext = false;
const pendingResolvers = [];
const originalSnapshots = [];

function makeQuery(collection, canonical) {
  return {
    _query: {
      path: { canonicalString: () => collection },
      canonicalId: () => canonical
    }
  };
}

function resolveNext() {
  const next = pendingResolvers.shift();
  assert.equal(typeof next, "function", "Deve esistere una richiesta simulata in attesa");
  next();
}

global.window = {
  addEventListener() {},
  runFirestoreGetWithRetry(query) {
    calls += 1;
    const callNumber = calls;
    if (rejectNext) {
      rejectNext = false;
      return Promise.reject(new Error("errore simulato"));
    }
    return new Promise((resolve) => {
      pendingResolvers.push(() => {
        const snapshot = { marker: `snapshot-${callNumber}`, query, docs: [] };
        originalSnapshots.push(snapshot);
        resolve(snapshot);
      });
    });
  }
};

async function run() {
  require(path.resolve(__dirname, "..", "firestore-inflight-read-coalescer.js"));
  window.HeraFirestoreInflightReadCoalescer.refreshInstallation();

  const queryA1 = makeQuery("personale", "personale|ob:createdAtasc");
  const queryA2 = makeQuery("personale", "personale|ob:createdAtasc");
  const options1 = { label: "PRIMO", timeoutMs: 9000, retries: 2 };
  const options2 = { label: "SECONDO", timeoutMs: 9000, retries: 2 };

  const first = window.runFirestoreGetWithRetry(queryA1, options1);
  const second = window.runFirestoreGetWithRetry(queryA2, options2);
  const third = window.runFirestoreGetWithRetry(queryA1, options1);
  assert.equal(calls, 1, "Tre richieste simultanee identiche devono avviare una sola lettura reale");
  resolveNext();
  const [a, b, c] = await Promise.all([first, second, third]);
  assert.equal(a, b, "Tutti i chiamanti devono ricevere lo stesso snapshot Firestore originale");
  assert.equal(b, c);
  assert.equal(a, originalSnapshots[0]);

  const next = window.runFirestoreGetWithRetry(queryA1, options1);
  assert.equal(calls, 2, "Dopo il completamento una nuova richiesta deve leggere nuovamente Firestore");
  resolveNext();
  const nextSnapshot = await next;
  assert.notEqual(nextSnapshot, a);

  const mezzi = window.runFirestoreGetWithRetry(makeQuery("mezzi", "mezzi|ob:createdAtasc"), options1);
  const personaleDifferent = window.runFirestoreGetWithRetry(makeQuery("personale", "personale|where:attivo"), options1);
  assert.equal(calls, 4, "Query differenti o collezioni differenti non devono essere condivise");
  resolveNext();
  resolveNext();
  await Promise.all([mezzi, personaleDifferent]);

  const squadreQuery = makeQuery("squadre", "squadre|all");
  const squadreA = window.runFirestoreGetWithRetry(squadreQuery, options1);
  const squadreB = window.runFirestoreGetWithRetry(squadreQuery, options1);
  assert.equal(calls, 6, "Squadre deve restare completamente esclusa dall'ottimizzazione");
  resolveNext();
  resolveNext();
  await Promise.all([squadreA, squadreB]);

  rejectNext = true;
  await assert.rejects(
    window.runFirestoreGetWithRetry(queryA1, options1),
    /errore simulato/,
    "Gli errori originali devono essere propagati"
  );
  const retry = window.runFirestoreGetWithRetry(queryA1, options1);
  assert.equal(calls, 8, "Dopo un errore la richiesta deve poter ripartire normalmente");
  resolveNext();
  await retry;

  assert.equal(pendingResolvers.length, 0, "Nessuna richiesta simulata deve restare sospesa");
  const state = window.HeraFirestoreInflightReadCoalescer.getState();
  assert.equal(state.inFlight, 0);
  assert.equal(state.stats.duplicateCallsShared, 2);
  assert.equal(state.stats.networkRequestsStarted, 6);
  assert.equal(state.stats.rejected, 1);

  console.log("✅ Snapshot Firestore originali mantenuti senza ricostruzione.");
  console.log("✅ Solo richieste simultanee identiche di personale/mezzi condivise.");
  console.log("✅ Squadre, query diverse, richieste successive ed errori restano indipendenti.");
}

run().catch((error) => {
  console.error("❌ Test condivisione letture Firestore non superato.");
  console.error(error);
  process.exit(1);
});

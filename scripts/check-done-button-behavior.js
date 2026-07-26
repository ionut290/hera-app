#!/usr/bin/env node
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const appSource = fs.readFileSync("app.js", "utf8");
const nativeSource = fs.readFileSync("native-android-runtime.js", "utf8");
const immediateSource = fs.readFileSync("fatto-button-immediate.js", "utf8");

function extractFunction(name) {
  const signatures = [`async function ${name}(`, `function ${name}(`];
  let start = -1;
  for (const signature of signatures) {
    start = appSource.indexOf(signature);
    if (start >= 0) break;
  }
  assert.notEqual(start, -1, `Funzione ${name} non trovata`);
  const signatureEnd = appSource.indexOf(") {", start);
  const bodyStart = appSource.indexOf("{", signatureEnd);
  let depth = 0;
  for (let index = bodyStart; index < appSource.length; index += 1) {
    if (appSource[index] === "{") depth += 1;
    if (appSource[index] === "}") depth -= 1;
    if (depth === 0) return appSource.slice(start, index + 1);
  }
  throw new Error(`Funzione ${name} non chiusa`);
}

function createContext(values = {}) {
  const context = { console, ...values };
  vm.createContext(context);
  return context;
}

function loadFunctions(context, names) {
  vm.runInContext(names.map(extractFunction).join("\n"), context);
}

async function main() {
  const groupedContext = createContext({
    buildImpiantoKey: (row) => `sap:${row.idSap}`,
    getPlantPriceCodeMqMap: () => ({}),
    mergePriceCodeMqMap: () => ({}),
    mergeMultiValue: (first, second) => [first, second].filter(Boolean).join(" | "),
    hasOrdinario: () => false,
    hasStraordinario: () => false,
    classifyTipoManutenzione: () => "Non specificata"
  });
  loadFunctions(groupedContext, [
    "firestoreDateToMillis",
    "isImpiantoDoneState",
    "combineImpiantiForView"
  ]);
  const groupedTodo = groupedContext.combineImpiantiForView([
    { id: "a", idSap: "100", done: false, doneAt: null, resetAt: null },
    { id: "b", idSap: "100", done: false, doneAt: null, resetAt: null }
  ]);
  assert.equal(groupedTodo.length, 1);
  assert.equal(groupedTodo[0].done, false, "Due righe DA FARE non devono diventare FATTO");
  assert.deepEqual(Array.from(groupedTodo[0].sourceIds), ["a", "b"]);

  const batchWrites = [];
  let batchCommits = 0;
  const batchContext = createContext({
    auth: { currentUser: { uid: "u1", email: "op@example.com", displayName: "Operatore" } },
    firebase: {
      firestore: {
        Timestamp: { fromDate: (date) => ({ millis: date.getTime() }) }
      }
    },
    getCommesseCollectionName: () => "commesse",
    db: {
      collection: () => ({
        doc: () => ({
          collection: () => ({
            doc: (id) => ({ id })
          })
        })
      }),
      batch: () => ({
        set: (ref, payload, options) => batchWrites.push({ ref, payload, options }),
        commit: async () => { batchCommits += 1; }
      })
    }
  });
  loadFunctions(batchContext, ["setImpiantoDone"]);
  await batchContext.setImpiantoDone("c1", ["a", "b"], true, {
    doneAt: new Date("2026-07-26T10:00:00Z")
  });
  assert.equal(batchWrites.length, 2, "Tutti i documenti raggruppati devono entrare nel batch");
  assert.equal(batchCommits, 1, "Il FATTO deve usare un solo commit atomico");
  assert.ok(batchWrites.every((entry) => entry.payload.done === true));

  batchContext.auth.currentUser = null;
  await assert.rejects(
    () => batchContext.setImpiantoDone("c1", ["a"], true),
    /Sessione scaduta/,
    "Una sessione assente non deve simulare un FATTO riuscito"
  );

  const persistedContext = createContext({
    selectedCommessaId: "c1",
    getImpiantoDocIds: () => ["a", "b"],
    getCommesseCollectionName: () => "commesse",
    isImpiantoDoneState: (data) => data.done === true,
    db: {
      collection: () => ({
        doc: () => ({
          collection: () => ({
            doc: (id) => ({
              get: async () => ({ exists: true, data: () => ({ done: id === "a" }) })
            })
          })
        })
      })
    }
  });
  loadFunctions(persistedContext, ["isImpiantoPersistedAsDone"]);
  assert.equal(
    await persistedContext.isImpiantoPersistedAsDone({ id: "a" }),
    false,
    "Un solo documento salvato non deve confermare l'intero impianto"
  );

  const waitingContext = createContext();
  loadFunctions(waitingContext, ["isActionWaitingForSync"]);
  assert.equal(waitingContext.isActionWaitingForSync(null), false);
  assert.equal(waitingContext.isActionWaitingForSync({ status: "pending" }), true);

  const whatsappHandler = extractFunction("handleImpiantoWhatsAppClick");
  assert.ok(
    whatsappHandler.indexOf("getCachedFattoPositionDecision(impianto)")
      < whatsappHandler.indexOf("recordFattoVisualEvidence(impianto, doneAt, doneBy)"),
    "GPS e distanza devono essere verificati prima della prova visiva"
  );
  assert.match(whatsappHandler, /const opened = openWhatsApp\(/);
  assert.match(whatsappHandler, /requireFirestoreConfirmation:\s*false/);
  assert.match(whatsappHandler, /queued_offline/);

  let deniedEvidenceWrites = 0;
  const deniedContext = createContext({
    auth: { currentUser: { uid: "u1" } },
    getWhazzupProcessingKey: () => "c1:sap:100",
    isImpiantoWhazzupProcessing: () => false,
    whazzupProcessingByImpianto: new Set(),
    getCachedFattoPositionDecision: () => ({ allowed: false, reason: "distance" }),
    closeDeferredWhatsAppTargetWindow: () => {},
    notifyFattoPositionDenied: () => {},
    clearImpiantoWhazzupProcessing: () => {},
    renderImpianti: () => {},
    recordFattoVisualEvidence: () => { deniedEvidenceWrites += 1; }
  });
  vm.runInContext(whatsappHandler, deniedContext);
  assert.equal(await deniedContext.handleImpiantoWhatsAppClick({ id: "a" }), false);
  assert.equal(deniedEvidenceWrites, 0, "Un FATTO oltre 4 km non deve creare prova visiva");

  const forceHandler = extractFunction("forceMarkDone");
  assert.match(forceHandler, /markImpiantoDone\(impianto,\s*\{\s*source:\s*"force"/);
  assert.doesNotMatch(forceHandler, /if\s*\(!isNetworkOffline\(\)\)/);

  assert.match(immediateSource, /addEventListener\("pointerup"/);
  assert.doesNotMatch(immediateSource, /addEventListener\("pointerdown"/);
  assert.match(nativeSource, /Geolocation\.watchPosition/);
  assert.match(nativeSource, /Geolocation\.clearWatch/);
  assert.match(nativeSource, /window\.dispatchEvent\(new CustomEvent\("hera:native-location"/);

  console.log("✅ Test comportamentali FATTO/WHAZZUP completati.");
}

main().catch((error) => {
  console.error("❌ Test comportamentali FATTO/WHAZZUP falliti:", error);
  process.exitCode = 1;
});

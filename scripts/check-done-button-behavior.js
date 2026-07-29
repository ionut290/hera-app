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
        Timestamp: { fromDate: (date) => ({ millis: date.getTime() }) },
        FieldValue: { serverTimestamp: () => ({ serverTimestamp: true }) }
      }
    },
    getCommesseCollectionName: () => "commesse",
    db: {
      collection: () => ({
        doc: () => ({
          collection: (name) => ({
            doc: (id) => ({ id, collection: name }),
            where: () => ({ get: async () => ({ docs: [] }) })
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
    doneAt: vm.runInContext('new Date("2026-07-26T10:00:00Z")', batchContext)
  });
  assert.equal(batchWrites.length, 2, "Tutti i documenti raggruppati devono entrare nel batch");
  assert.equal(batchCommits, 1, "Il FATTO deve usare un solo commit atomico");
  assert.ok(batchWrites.slice(0, 2).every((entry) => entry.payload.done === true));
  assert.ok(batchWrites.slice(0, 2).every((entry) => entry.payload.dataEsecuzione === "2026-07-26"));
  assert.ok(batchWrites.slice(0, 2).every((entry) => entry.payload.oraEsecuzione === "10:00"));

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
    whatsappHandler.indexOf("validateImpiantoCoordinates(impianto)")
      < whatsappHandler.indexOf("recordFattoVisualEvidence(impianto, doneAt, doneBy)"),
    "Le coordinate impianto devono essere validate prima della prova visiva"
  );
  assert.doesNotMatch(whatsappHandler, /getCachedFattoPositionDecision|refreshFattoPositionDecision|getCurrentPositionOnce|currentUserPos|distanceFromUser/);
  assert.match(whatsappHandler, /const opened = openWhatsApp\(/);
  assert.ok(
    whatsappHandler.indexOf("recordFattoVisualEvidence(impianto, doneAt, doneBy)")
      < whatsappHandler.indexOf("const opened = openWhatsApp("),
    "Lo stato FATTO deve essere salvato prima di WhatsApp"
  );
  assert.ok(
    whatsappHandler.indexOf("const opened = openWhatsApp(")
      < whatsappHandler.indexOf("markImpiantoDoneVisualFallback(impianto, { doneAt, doneBy });"),
    "Il trasferimento nei FATTI deve iniziare dopo WhatsApp"
  );
  assert.match(whatsappHandler, /requireFirestoreConfirmation:\s*false/);
  assert.match(whatsappHandler, /queued_offline/);

  const coordinatesContext = createContext();
  loadFunctions(coordinatesContext, ["validateImpiantoCoordinates"]);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsY: "44.50", gpsX: "11.34" }).valid, true);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsY: "", gpsX: "11.34" }).valid, false);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsX: "11.34" }).valid, false);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsY: "nord", gpsX: "est" }).valid, false);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsY: 91, gpsX: 11 }).valid, false);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsY: 44, gpsX: -181 }).valid, false);

  let evidenceWrites = 0;
  let whatsappOpens = 0;
  let doneMoves = 0;
  const flowContext = createContext({
    auth: { currentUser: { uid: "u1", displayName: "Operatore" } },
    selectedCommessaId: "c1",
    validateImpiantoCoordinates: coordinatesContext.validateImpiantoCoordinates,
    notifyInvalidImpiantoCoordinates: () => {},
    getWhazzupProcessingKey: () => "c1:sap:100",
    isImpiantoWhazzupProcessing: () => false,
    whazzupProcessingByImpianto: new Set(),
    openDeferredWhatsAppTargetWindow: () => null,
    closeDeferredWhatsAppTargetWindow: () => {},
    clearImpiantoWhazzupProcessing: () => {},
    setImpiantoFattoSavingState: () => {},
    isNetworkOffline: () => false,
    recordFattoVisualEvidence: async () => { evidenceWrites += 1; return true; },
    cacheFattoVisualEvidence: () => {},
    markWhazzupSafetyPressed: () => {},
    upsertWhazzupPendingDoneEntry: () => {},
    renderImpianti: () => {},
    openWhatsApp: () => { whatsappOpens += 1; return true; },
    markImpiantoDoneVisualFallback: () => {},
    setImpiantiViewMode: () => {},
    auditLogWhazzupClick: async () => null,
    forceMoveImpiantoToFatti: async () => { doneMoves += 1; return true; },
    updateAuditLogWhazzupClick: async () => {},
    getPendingActionForImpianto: () => ({ status: "pending" }),
    isActionWaitingForSync: () => true,
    markPendingActionStatus: () => {},
    updateConnectivityStatus: () => {},
    markImpiantoDoneRecoveryRequired: async () => {},
    alert: () => {}
  });
  vm.runInContext(whatsappHandler, flowContext);
  assert.equal(await flowContext.handleImpiantoWhatsAppClick({ id: "a", gpsY: "44.50", gpsX: "11.34" }), true);
  assert.equal(evidenceWrites, 1, "Coordinate valide conservano il salvataggio FATTO");
  assert.equal(whatsappOpens, 1, "Coordinate valide conservano l’apertura WhatsApp");
  assert.equal(doneMoves, 1, "Coordinate valide conservano il trasferimento nei FATTI");

  assert.equal(await flowContext.handleImpiantoWhatsAppClick({ id: "b", gpsY: "", gpsX: "11.34" }), false);
  assert.equal(evidenceWrites, 1, "Coordinate mancanti bloccano prima del salvataggio FATTO");
  assert.equal(whatsappOpens, 1, "Coordinate mancanti bloccano prima di WhatsApp");

  let completedWhatsappOpens = 0;
  const completedContext = createContext({
    auth: { currentUser: { uid: "u1", displayName: "Operatore" } },
    getCurrentWhatsAppOperatorName: () => "Operatore",
    openWhatsApp: () => { completedWhatsappOpens += 1; return true; },
    alert: () => {}
  });
  loadFunctions(completedContext, ["handleCompletedImpiantoWhatsAppClick"]);
  assert.equal(
    completedContext.handleCompletedImpiantoWhatsAppClick({ done: true, doneAt: "2026-07-29T10:00:00Z" }),
    true,
    "FATTO DAL deve riaprire Whazzup"
  );
  assert.equal(completedWhatsappOpens, 1);
  const completedHandler = extractFunction("handleCompletedImpiantoWhatsAppClick");
  assert.doesNotMatch(completedHandler, /validateImpiantoCoordinates|getCurrentPositionOnce|currentUserPos|distanceFromUser/);

  const forceHandler = extractFunction("forceMarkDone");
  assert.match(forceHandler, /markImpiantoDone\(impianto,\s*\{\s*source:\s*"force"/);
  assert.doesNotMatch(forceHandler, /if\s*\(!isNetworkOffline\(\)\)/);

  assert.doesNotMatch(immediateSource, /addEventListener\("pointerup"/);
  assert.doesNotMatch(immediateSource, /addEventListener\("click"/);
  assert.match(nativeSource, /Geolocation\.watchPosition/);
  assert.match(nativeSource, /Geolocation\.clearWatch/);
  assert.match(nativeSource, /window\.dispatchEvent\(new CustomEvent\("hera:native-location"/);

  console.log("✅ Test comportamentali FATTO/WHAZZUP completati.");
}

main().catch((error) => {
  console.error("❌ Test comportamentali FATTO/WHAZZUP falliti:", error);
  process.exitCode = 1;
});

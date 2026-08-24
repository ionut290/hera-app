#!/usr/bin/env node
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const appSource = fs.readFileSync("app.js", "utf8");
const photoOrderSource = fs.readFileSync("android-whazzup-photo-order.js", "utf8");
const styleSource = fs.readFileSync("style.css", "utf8");
const nativeSource = fs.readFileSync("native-android-runtime.js", "utf8");
const immediateSource = fs.readFileSync("fatto-button-immediate.js", "utf8");
const touchGuardSource = fs.readFileSync("fatto-scroll-guard.js", "utf8");
const packageJson = JSON.parse(fs.readFileSync("package.json", "utf8"));

function extractFunctionFrom(source, name) {
  const signatures = [`async function ${name}(`, `function ${name}(`];
  let start = -1;
  for (const signature of signatures) {
    start = source.indexOf(signature);
    if (start >= 0) break;
  }
  assert.notEqual(start, -1, `Funzione ${name} non trovata`);
  const signatureEnd = source.indexOf(") {", start);
  const bodyStart = source.indexOf("{", signatureEnd);
  let depth = 0;
  for (let index = bodyStart; index < source.length; index += 1) {
    if (source[index] === "{") depth += 1;
    if (source[index] === "}") depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }
  throw new Error(`Funzione ${name} non chiusa`);
}

function extractFunction(name) {
  return extractFunctionFrom(appSource, name);
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
  const maintenanceContext = createContext();
  loadFunctions(maintenanceContext, ["splitCodes", "hasOrdinario", "hasStraordinario", "classifyTipoManutenzione"]);
  assert.equal(maintenanceContext.hasOrdinario("A1"), false, "Il ripristino del 19/08 non tratta A1 come ordinario");
  assert.equal(maintenanceContext.hasStraordinario("A1"), true, "Il ripristino del 19/08 tratta A1 come straordinario");
  assert.equal(maintenanceContext.classifyTipoManutenzione("A1"), "Straordinaria");

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
  const batchDoneAt = vm.runInContext('new Date("2026-07-26T10:00:00Z")', batchContext);
  await batchContext.setImpiantoDone("c1", ["a", "b"], true, {
    doneAt: batchDoneAt
  });
  assert.equal(batchWrites.length, 2, "Tutti i documenti raggruppati devono entrare nel batch");
  assert.equal(batchCommits, 1, "Il FATTO deve usare un solo commit atomico");
  assert.ok(batchWrites.slice(0, 2).every((entry) => entry.payload.done === true));
  assert.ok(batchWrites.slice(0, 2).every((entry) => entry.payload.dataEsecuzione === "2026-07-26"));
  const expectedExecutionTime = [
    String(batchDoneAt.getHours()).padStart(2, "0"),
    String(batchDoneAt.getMinutes()).padStart(2, "0")
  ].join(":");
  assert.ok(batchWrites.slice(0, 2).every((entry) => entry.payload.oraEsecuzione === expectedExecutionTime));

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
  assert.match(whatsappHandler, /const opened = hasWhazzupPhotos/);
  assert.match(whatsappHandler, /shareWhazzupWithPhotos/);
  assert.match(whatsappHandler, /openWhatsApp\(/);
  assert.ok(
    whatsappHandler.indexOf("markImpiantoDoneVisualFallback(impianto, { doneAt, doneBy });")
      < whatsappHandler.indexOf("recordFattoVisualEvidence(impianto, doneAt, doneBy)"),
    "Il click deve spostare subito la scheda prima delle scritture accessorie"
  );
  assert.doesNotMatch(whatsappHandler, /await recordFattoVisualEvidence/);
  assert.doesNotMatch(
    whatsappHandler,
    /if \(!evidenceSaved && !isNetworkOffline\(\)\) \{\s*closeDeferredWhatsAppTargetWindow\([\s\S]*?return false;/,
    "Un errore della prova visiva non deve bloccare WhatsApp"
  );
  assert.ok(
    whatsappHandler.indexOf("markImpiantoDoneVisualFallback(impianto, { doneAt, doneBy });")
      < whatsappHandler.indexOf("const opened = hasWhazzupPhotos"),
    "Il trasferimento visivo nei FATTI deve precedere WhatsApp"
  );
  assert.match(whatsappHandler, /requireFirestoreConfirmation:\s*false/);
  assert.match(whatsappHandler, /queued_offline/);

  assert.doesNotMatch(
    appSource,
    /has-completion-evidence|Già premuto FATTO/,
    "Una prova FATTO pregressa non deve rendere giallo un impianto ancora Da fare"
  );

  const coordinatesContext = createContext();
  loadFunctions(coordinatesContext, ["validateImpiantoCoordinates"]);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsY: "44.50", gpsX: "11.34" }).valid, true);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsY: "", gpsX: "11.34" }).valid, false);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsX: "11.34" }).valid, false);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsY: "nord", gpsX: "est" }).valid, false);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsY: 91, gpsX: 11 }).valid, false);
  assert.equal(coordinatesContext.validateImpiantoCoordinates({ gpsY: 44, gpsX: -181 }).valid, false);

  let evidenceWrites = 0;
  let evidenceShouldSave = true;
  let whatsappOpens = 0;
  let doneMoves = 0;
  let accessorySaveFailures = 0;
  const flowContext = createContext({
    auth: { currentUser: { uid: "u1", displayName: "Operatore" } },
    selectedCommessaId: "c1",
    validateImpiantoCoordinates: coordinatesContext.validateImpiantoCoordinates,
    notifyInvalidImpiantoCoordinates: () => {},
    getWhazzupProcessingKey: (impianto) => `c1:${impianto.id}`,
    isImpiantoWhazzupProcessing: () => false,
    whazzupProcessingByImpianto: new Set(),
    openDeferredWhatsAppTargetWindow: () => null,
    closeDeferredWhatsAppTargetWindow: () => {},
    clearImpiantoWhazzupProcessing: (impianto) => flowContext.whazzupProcessingByImpianto.delete(`c1:${impianto.id}`),
    setImpiantoFattoSavingState: () => {},
    isNetworkOffline: () => false,
    recordFattoVisualEvidence: async () => { evidenceWrites += 1; return evidenceShouldSave; },
    cacheFattoVisualEvidence: () => {},
    markWhazzupSafetyPressed: () => {},
    upsertWhazzupPendingDoneEntry: () => {},
    renderImpianti: () => {},
    getWhazzupPhotos: () => [],
    shareWhazzupWithPhotos: async () => true,
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
    handleImpiantoDoneSaveFailure: async () => { accessorySaveFailures += 1; },
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

  evidenceShouldSave = false;
  assert.equal(await flowContext.handleImpiantoWhatsAppClick({ id: "c", gpsY: "44.50", gpsX: "11.34" }), true);
  assert.equal(evidenceWrites, 2, "Il tentativo della prova visiva viene mantenuto");
  assert.equal(whatsappOpens, 2, "WhatsApp deve aprirsi anche se la prova visiva non si salva");
  assert.equal(doneMoves, 2, "L'impianto deve passare nei FATTI anche se la prova visiva non si salva");
  assert.equal(accessorySaveFailures, 1, "L'errore accessorio viene segnalato senza bloccare il flusso");

  flowContext.recordFattoVisualEvidence = () => new Promise(() => {});
  const nonBlockingResult = await Promise.race([
    flowContext.handleImpiantoWhatsAppClick({ id: "d", gpsY: "44.50", gpsX: "11.34" }),
    new Promise((resolve) => setTimeout(() => resolve("timeout"), 150))
  ]);
  assert.equal(nonBlockingResult, true, "Una scrittura accessoria sospesa non deve bloccare FATTO");
  assert.equal(whatsappOpens, 3, "WhatsApp deve aprirsi anche con la prova visiva sospesa");
  assert.equal(doneMoves, 3, "Il salvataggio reale deve partire anche con la prova visiva sospesa");

  let resetWrites = 0;
  let resetLocalState = null;
  const resetContext = createContext({
    selectedCommessaId: "c1",
    selectedCommessaName: "Commessa prova",
    currentUser: { displayName: "Operatore" },
    getImpiantoDocIds: () => ["a"],
    canUseImpiantoAction: () => true,
    isNetworkOffline: () => false,
    retrySetImpiantoDone: async (_commessaId, _ids, done) => { resetWrites += 1; return done === false; },
    clearImpiantoWhazzupProcessing: () => {},
    getWhazzupSafetyState: () => null,
    clearWhazzupPendingDoneEntry: () => {},
    removePendingDoneActionsForImpianto: () => {},
    trackLocalSheetMutation: () => {},
    updateImpiantoLocalState: (_ids, patch) => { resetLocalState = patch; },
    buildImpiantoKey: () => "id:a",
    clearActionUsed: () => {},
    updateConnectivityStatus: () => {},
    renderImpianti: () => {},
    scheduleCommessaSheetSync: () => {},
    alert: () => {}
  });
  loadFunctions(resetContext, ["resetImpianto"]);
  assert.equal(await resetContext.resetImpianto({ id: "a", done: true }), true, "RESET deve completarsi con un click");
  assert.equal(resetWrites, 1, "RESET deve avviare una sola scrittura");
  assert.equal(resetLocalState.done, false, "RESET deve riportare localmente l'impianto in Da fare");

  let completedWhatsappOpens = 0;
  const completedContext = createContext({
    auth: { currentUser: { uid: "u1", displayName: "Operatore" } },
    getCurrentWhatsAppOperatorName: () => "Operatore",
    getWhazzupPhotos: () => [],
    shareWhazzupWithPhotos: async () => true,
    openWhatsApp: () => { completedWhatsappOpens += 1; return true; },
    alert: () => {}
  });
  loadFunctions(completedContext, ["handleCompletedImpiantoWhatsAppClick"]);
  assert.equal(
    await completedContext.handleCompletedImpiantoWhatsAppClick({ done: true, doneAt: "2026-07-29T10:00:00Z" }),
    true,
    "FATTO DAL deve riaprire Whazzup"
  );
  assert.equal(completedWhatsappOpens, 1);
  const completedHandler = extractFunction("handleCompletedImpiantoWhatsAppClick");
  assert.doesNotMatch(completedHandler, /validateImpiantoCoordinates|getCurrentPositionOnce|currentUserPos|distanceFromUser/);

  assert.match(appSource, /WHAZZUP_PHOTO_MAX_AGE_MS\s*=\s*10\s*\*\s*60\s*\*\s*60\s*\*\s*1000/);
  assert.match(appSource, /indexedDB\.open\(WHAZZUP_PHOTO_DB_NAME,\s*1\)/);
  assert.match(appSource, /void restorePersistedWhazzupPhotos\(\)/);
  assert.match(appSource, /await persistWhazzupPhotos\(key, validFiles, savedAt, normalizedNotes\)/);
  assert.match(appSource, /notes:\s*normalizeWhazzupPhotoNotes\(notes, files\.length\)/);
  assert.match(appSource, /function openWhazzupPhotoManager\(impianto, button\)/);
  assert.match(appSource, /data-photo-action="view"/);
  assert.match(appSource, /data-photo-action="replace"/);
  assert.match(appSource, /data-photo-action="delete"/);
  assert.match(appSource, /data-manager-action="add"/);
  assert.match(appSource, /data-manager-action="replace-all"/);
  assert.match(appSource, /data-manager-action="delete-all"/);
  assert.match(appSource, /data-manager-action="done">✅ \$\{submitLabel\}/);
  const photoManagerHandler = extractFunction("openWhazzupPhotoManager");
  assert.match(photoManagerHandler, /await handleImpiantoWhatsAppClick\(impianto\)/);
  assert.match(photoManagerHandler, /await handleCompletedImpiantoWhatsAppClick\(impianto\)/);
  assert.match(photoManagerHandler, /isImpiantoWhazzupProcessing\(impianto\)/);
  assert.doesNotMatch(photoManagerHandler, /markImpiantoDone\(|forceMoveImpiantoToFatti\(|setImpiantoDone\(/);
  assert.match(styleSource, /\.impianto-primary-actions\s*\{[\s\S]*grid-template-columns:\s*repeat\(3, minmax\(0, 1fr\)\)/);
  assert.match(styleSource, /data-action-key="whatsapp-attachment"\][\s\S]*order:\s*2/);
  assert.match(styleSource, /\.whazzup-photo-manager-actions \.whazzup-photo-manager-submit\s*\{/);
  assert.match(appSource, /function buildOrderedWhazzupShareFiles\(files\)/);
  assert.match(appSource, /Foto-\$\{String\(index \+ 1\)\.padStart\(2, "0"\)\}/);
  assert.ok(packageJson.dependencies["@capacitor/filesystem"], "Manca il filesystem nativo Android");
  assert.ok(packageJson.dependencies["@capacitor/share"], "Manca la condivisione nativa Android");
  assert.match(appSource, /function getNativeAndroidWhazzupSharePlugins\(\)/);
  assert.match(appSource, /plugins\.filesystem\.writeFile\(\{/);
  assert.match(appSource, /directory: "CACHE"/);
  const nativePhotoShareHandler = extractFunctionFrom(photoOrderSource, "shareWhazzupPhotosNativeAndroidInOrder");
  const photoShareHandler = extractFunction("shareWhazzupWithPhotos");
  assert.match(nativePhotoShareHandler, /files: fileUris/);
  assert.doesNotMatch(nativePhotoShareHandler, /text:\s*message/);
  assert.ok(
    nativePhotoShareHandler.indexOf("safeOpenWhatsAppMessage(message)")
      < nativePhotoShareHandler.indexOf("await sharePhotosThroughDedicatedPlugin"),
    "Il messaggio Whazzup deve essere aperto prima delle foto"
  );
  assert.match(photoShareHandler, /shareWhazzupPhotosNativeAndroid\(orderedFiles, message\)/);
  assert.match(photoShareHandler, /files: orderedFiles[\s\S]*text: message/);
  assert.match(appSource, /await deletePersistedWhazzupPhotos\(getWhazzupPhotoKey\(impianto\)\)/);

  const payloadContext = createContext({
    selectedCommessaId: "c1",
    selectedCommessaName: "Commessa prova",
    auth: { currentUser: { uid: "u1", displayName: "Operatore" } },
    currentUser: null,
    impiantoWhatsAppTemplateCache: new Map(),
    getCommessaNoteLinkedNotes: () => [],
    getCommessaNoteTitle: () => "",
    buildImpiantoKey: (impianto) => `id:${impianto.id || "mancante"}`,
    hasOrdinario: () => false,
    hasStraordinario: () => false
  });
  loadFunctions(payloadContext, [
    "firestoreDateToMillis",
    "formatDoneDateTime",
    "getImpiantoWhatsAppTemplateCacheKey",
    "getCurrentWhatsAppOperatorName",
    "getDeviceWhatsAppDateLabel",
    "getImpiantoWhatsAppTemplateSignature",
    "buildImpiantoWhatsAppTemplate",
    "prepareImpiantoWhatsAppTemplate",
    "buildImpiantoWhatsAppPayload"
  ]);
  const completePayload = payloadContext.buildImpiantoWhatsAppPayload({
    id: "a",
    idSap: "SAP-100",
    denominazione: "Impianto prova",
    comune: "Bologna"
  }, { doneAt: "2026-07-30T10:15:00Z" });
  assert.match(completePayload.message, /🟢 IMPIANTO FATTO/);
  assert.match(completePayload.message, /Impianto prova/);
  assert.ok(completePayload.message.trim(), "Il messaggio Whazzup completo non deve essere vuoto");
  assert.match(completePayload.appUrl, /^whatsapp:\/\/send\?text=.+/);
  assert.match(completePayload.webUrl, /^https:\/\/wa\.me\/\?text=.+/);

  const fallbackPayload = payloadContext.buildImpiantoWhatsAppPayload(
    { id: "b" },
    { doneAt: "data-non-valida" }
  );
  assert.ok(fallbackPayload.message.trim(), "Dati secondari mancanti non devono produrre un messaggio vuoto");
  assert.match(fallbackPayload.message, /Impianto: -/);
  assert.match(decodeURIComponent(fallbackPayload.appUrl), /🟢 IMPIANTO FATTO/);

  const forceHandler = extractFunction("forceMarkDone");
  assert.match(forceHandler, /markImpiantoDone\(impianto,\s*\{\s*source:\s*"force"/);
  assert.doesNotMatch(forceHandler, /if\s*\(!isNetworkOffline\(\)\)/);

  assert.doesNotMatch(immediateSource, /document\.addEventListener\("pointerup"/);
  assert.doesNotMatch(immediateSource, /document\.addEventListener\("click"/);
  assert.doesNotMatch(touchGuardSource, /BLOCK_AFTER_SCROLL_MS|MOVE_THRESHOLD_PX|findFattoButton/);
  assert.doesNotMatch(touchGuardSource, /stopImmediatePropagation\(\)/, "Nessun guard deve mangiare il primo click FATTO");
  assert.doesNotMatch(immediateSource, /operation\s*=\s*await enqueue/);
  assert.match(immediateSource, /const operationTask = enqueue/);

  let wrappedClicks = 0;
  const hangingQueueContext = createContext({
    Blob,
    JSON,
    Date,
    Promise,
    setTimeout,
    clearTimeout,
    navigator: { onLine: true },
    indexedDB: { open: () => ({}) },
    CustomEvent: class {
      constructor(type, init = {}) { this.type = type; this.detail = init.detail; }
    },
    document: {
      visibilityState: "visible",
      documentElement: { dataset: {} },
      addEventListener: () => {},
      querySelector: () => null,
      createElement: () => ({ dataset: {} }),
      head: { appendChild: () => {} }
    },
    selectedCommessaId: "c1",
    auth: { currentUser: { uid: "u1" } },
    handleImpiantoWhatsAppClick: async () => { wrappedClicks += 1; return true; },
    openWhatsApp: () => true,
    addEventListener: () => {},
    dispatchEvent: () => {},
    setInterval: () => 0,
    clearInterval: () => {}
  });
  hangingQueueContext.window = hangingQueueContext;
  vm.runInContext(immediateSource, hangingQueueContext, { filename: "fatto-button-immediate.js" });
  const hangingQueueResult = await Promise.race([
    hangingQueueContext.handleImpiantoWhatsAppClick({ id: "queue-hang" }),
    new Promise((resolve) => setTimeout(() => resolve("timeout"), 150))
  ]);
  assert.equal(hangingQueueResult, true, "IndexedDB sospeso non deve bloccare il click FATTO");
  assert.equal(wrappedClicks, 1, "Il wrapper deve inoltrare il click una sola volta");
  assert.match(nativeSource, /Geolocation\.watchPosition/);
  assert.match(nativeSource, /Geolocation\.clearWatch/);
  assert.match(nativeSource, /window\.dispatchEvent\(new CustomEvent\("hera:native-location"/);

  console.log("✅ Test comportamentali FATTO/WHAZZUP completati.");
}

main().catch((error) => {
  console.error("❌ Test comportamentali FATTO/WHAZZUP falliti:", error);
  process.exitCode = 1;
});

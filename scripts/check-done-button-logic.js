#!/usr/bin/env node
const fs = require('fs');
const path = require('path');

const appPath = path.join(__dirname, '..', 'app.js');
const source = fs.readFileSync(appPath, 'utf8');

function fail(message) {
  console.error(`❌ ${message}`);
  process.exitCode = 1;
}

function pass(message) {
  console.log(`✅ ${message}`);
}

function extractFunction(name) {
  const asyncSignature = `async function ${name}(`;
  const functionSignature = `function ${name}(`;
  let signature = asyncSignature;
  let start = source.indexOf(asyncSignature);
  if (start < 0) {
    signature = functionSignature;
    start = source.indexOf(functionSignature);
  }
  if (start < 0) return '';
  const signatureEnd = source.indexOf(') {', start);
  const bodyStart = signatureEnd >= 0 ? source.indexOf('{', signatureEnd) : source.indexOf('{', start);
  let depth = 0;
  for (let i = bodyStart; i < source.length; i += 1) {
    if (source[i] === '{') depth += 1;
    if (source[i] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, i + 1);
  }
  return '';
}

function requireIncludes(scope, needle, description) {
  if (!scope.includes(needle)) fail(description);
  else pass(description);
}

const criticalComment = 'LOGICA CRITICA PULSANTE FATTO - NON MODIFICARE SENZA TEST';
requireIncludes(source, criticalComment, 'commento di protezione presente nel codice');

const markDone = extractFunction('markImpiantoDone');
if (!markDone) fail('funzione markImpiantoDone presente');
else {
  pass('funzione markImpiantoDone presente');
  requireIncludes(markDone, 'validateImpiantoCoordinates(impianto)', 'Fatto valida le coordinate salvate nella scheda impianto');
  requireIncludes(markDone, 'notifyInvalidImpiantoCoordinates()', 'Fatto blocca con un messaggio le coordinate impianto non valide');
  requireIncludes(markDone, 'isNetworkOffline()', 'Fatto mantiene il ramo offline');
  requireIncludes(markDone, 'upsertPendingDoneAction', 'Fatto offline conserva azione pending WhatsApp');
  requireIncludes(markDone, 'setImpiantoDone(selectedCommessaId, ids, true', 'Fatto salva su Firebase con stato true');
  requireIncludes(markDone, 'updateImpiantoLocalState(ids, {\n      done: true', 'Fatto aggiorna lo stato locale a done true');
  requireIncludes(markDone, 'setImpiantiViewMode("done")', 'Fatto continua a spostare la UI sulla lista Fatti');
  requireIncludes(markDone, 'retrySetImpiantoDone', 'Fatto mantiene il retry Firebase');
  requireIncludes(markDone, 'queueSheetExportForAdmin', 'Fatto operatori conserva export/admin queue');
  requireIncludes(markDone, 'scheduleCommessaSheetSync', 'Fatto admin conserva sincronizzazione sheet');
  requireIncludes(markDone, 'publishGlobalNotificationEvent("impianto-done"', 'Fatto conserva notifica globale impianto-done');
}

const forceDone = extractFunction('canUseForceImpiantoDone');
if (!forceDone) fail('funzione canUseForceImpiantoDone presente');
else {
  pass('funzione canUseForceImpiantoDone presente');
  requireIncludes(forceDone, 'validateImpiantoCoordinates(impianto)', 'FORZA usa la stessa validazione coordinate di FATTO');
  if (forceDone.includes('hasFailedFattoAttemptForImpianto')) fail('FORZA non deve richiedere un precedente errore FATTO');
  else pass('FORZA non richiede un precedente errore FATTO');
}

const setDone = extractFunction('setImpiantoDone');
if (!setDone) fail('funzione setImpiantoDone presente');
else {
  pass('funzione setImpiantoDone presente');
  requireIncludes(setDone, 'firebase.firestore.Timestamp.fromDate(doneAtDate)', 'setImpiantoDone conserva timestamp Firebase doneAt');
  requireIncludes(setDone, 'doneBy: done ? (options.doneBy || user.displayName || user.email || "Operatore") : ""', 'setImpiantoDone conserva doneBy e reset a vuoto');
  requireIncludes(setDone, 'doneByUid: done ? String(options.doneByUid || user.uid || "") : ""', 'setImpiantoDone conserva doneByUid');
  requireIncludes(setDone, 'doneByEmail: done ? String(options.doneByEmail || user.email || "") : ""', 'setImpiantoDone conserva doneByEmail');
  requireIncludes(setDone, 'payload.resetAt = null', 'setImpiantoDone pulisce resetAt quando Fatto');
  requireIncludes(setDone, 'payload.resetBy = ""', 'setImpiantoDone pulisce resetBy quando Fatto');
  requireIncludes(setDone, 'const batch = db.batch()', 'setImpiantoDone usa una scrittura atomica');
  requireIncludes(setDone, 'await batch.commit()', 'setImpiantoDone conferma il batch atomico');
  requireIncludes(setDone, 'throw new Error("Sessione scaduta', 'setImpiantoDone non simula successo senza login');
}

const renderStart = source.indexOf('Questo è il pulsante operativo visibile');
const renderArea = source.slice(renderStart >= 0 ? renderStart : source.indexOf('if (!impianto.done) {'), source.indexOf('function openImpiantoEditor'));
requireIncludes(renderArea, '"whatsapp",\n        "✉️",\n        "Whazzup / Fatto"', 'render impianto conserva il pulsante operativo Whazzup / Fatto');
requireIncludes(renderArea, 'await handleImpiantoWhatsAppClick(impianto);', 'pulsante Whazzup / Fatto conserva handler WhatsApp esistente');
requireIncludes(renderArea, 'hiddenMoveDoneBtn.dataset.hiddenMoveDoneBtn = "1"', 'render impianto conserva pulsante nascosto di spostamento in Fatti');
requireIncludes(source, 'const forceDoneEnabled = forceDoneCoordinatesValid', 'FORZA dipende dalla validità delle coordinate impianto');
requireIncludes(renderArea, 'Sposta subito questo impianto nei FATTI', 'FORZA mostra che sposta subito nei FATTI');
requireIncludes(renderArea, 'await markImpiantoDone(impianto, { source: "whatsapp" });', 'pulsante nascosto continua a chiamare markImpiantoDone con source whatsapp');

const whatsappHandler = extractFunction('handleImpiantoWhatsAppClick');
if (!whatsappHandler) fail('funzione handleImpiantoWhatsAppClick presente');
else {
  pass('funzione handleImpiantoWhatsAppClick presente');
  if (/getCachedFattoPositionDecision|refreshFattoPositionDecision|getCurrentPositionOnce|currentUserPos|distanceFromUser/.test(whatsappHandler)) {
    fail('Fatto non usa posizione o distanza dell’operatore per autorizzare il flusso');
  } else {
    pass('Fatto non usa posizione o distanza dell’operatore per autorizzare il flusso');
  }
  const validationIndex = whatsappHandler.indexOf('validateImpiantoCoordinates(impianto)');
  const evidenceIndex = whatsappHandler.indexOf('recordFattoVisualEvidence(impianto, doneAt, doneBy)');
  const evidenceRenderIndex = whatsappHandler.indexOf('renderImpianti();', evidenceIndex);
  const openWhatsappIndex = whatsappHandler.indexOf('const opened = hasWhazzupPhotos');
  const localMoveIndex = whatsappHandler.indexOf('markImpiantoDoneVisualFallback(impianto, { doneAt, doneBy });');
  if (validationIndex >= 0 && evidenceIndex > validationIndex && evidenceRenderIndex > evidenceIndex && openWhatsappIndex > evidenceRenderIndex && localMoveIndex > openWhatsappIndex) {
    pass('Fatto salva e mostra lo stato, apre WhatsApp e solo dopo trasferisce nei FATTI');
  } else {
    fail('Ordine richiesto: salvataggio/visuale -> WhatsApp -> trasferimento FATTI');
  }
  const evidenceFailureStart = whatsappHandler.indexOf('if (!evidenceSaved && !isNetworkOffline())');
  const evidenceFailureEnd = whatsappHandler.indexOf('cacheFattoVisualEvidence(impianto, doneAt)', evidenceFailureStart);
  const evidenceFailureBlock = evidenceFailureStart >= 0 && evidenceFailureEnd > evidenceFailureStart
    ? whatsappHandler.slice(evidenceFailureStart, evidenceFailureEnd)
    : '';
  if (!evidenceFailureBlock) {
    fail('Fatto conserva la gestione dell’errore della prova visiva iniziale');
  } else if (/return false|closeDeferredWhatsAppTargetWindow/.test(evidenceFailureBlock)) {
    fail('Un errore di salvataggio iniziale non deve chiudere o bloccare WhatsApp');
  } else {
    pass('Un errore di salvataggio iniziale non chiude e non blocca WhatsApp');
  }
  requireIncludes(source, 'WhatsApp verrà aperto comunque', 'L’utente viene informato che WhatsApp continua anche con errore salvataggio');
  requireIncludes(whatsappHandler, 'setImpiantoFattoSavingState(impianto, true)', 'Fatto disabilita il pulsante durante il salvataggio');
  requireIncludes(whatsappHandler, 'if (!impianto || impianto.done) return false', 'Fatto ignora impianti già completati');
  requireIncludes(whatsappHandler, 'requireFirestoreConfirmation: false', 'Fatto conserva il salvataggio Firebase con coda offline');
  requireIncludes(whatsappHandler, 'skipImpiantoCoordinatesValidation: true', 'Fatto non ripete la validazione coordinate già superata');
  requireIncludes(whatsappHandler, 'queued_offline', 'Fatto registra l’esito accodato offline');
}

const persistedCheck = extractFunction('isImpiantoPersistedAsDone');
requireIncludes(persistedCheck, 'snapshots.every(', 'verifica FATTO richiede tutti i documenti raggruppati');

const backgroundSafetyCheck = extractFunction('runWhazzupPendingDoneSafetyCheck');
if (!backgroundSafetyCheck) fail('funzione runWhazzupPendingDoneSafetyCheck presente');
else {
  pass('funzione runWhazzupPendingDoneSafetyCheck presente');
  requireIncludes(backgroundSafetyCheck, 'if (!persisted && !isNetworkOffline())', 'rientro in app ritenta FATTO solo quando online');
  requireIncludes(backgroundSafetyCheck, 'await forceMoveImpiantoToFatti(impianto, { source: "whatsapp" })', 'rientro in app completa automaticamente il passaggio nei FATTI');
}

if (process.exitCode) process.exit(process.exitCode);
console.log('✅ Controlli logica pulsante Fatto completati.');

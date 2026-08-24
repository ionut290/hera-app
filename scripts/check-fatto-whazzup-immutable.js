#!/usr/bin/env node
"use strict";

const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const LOCK_MARKER = "BLOCCO_IRREVOCABILE_FATTO_WHAZZUP_V1";
const LOCKED_AT_COMMIT = "7bd2b6f0446a4cb64bcfbbb4e36fbf3ea83fdc67";
const LOCKED_FUNCTIONS = Object.freeze({
  "isImpiantoDoneState": "b55b62348bbe05cec7830a77a78e0eb11f245873",
  "combineImpiantiForView": "82df4ae3e1bee4e72b9665289326c38b33546018",
  "markImpiantoDone": "c5c463ba18cab7795a008a0c063aedb0ff3b0c2d",
  "markImpiantoDoneVisualFallback": "b0c4391c1ae68c4d7880859589cf1fe0f80b2113",
  "retrySetImpiantoDone": "49efcea672182a860e102c1ceabb2ed5579ef0c2",
  "resetImpianto": "7eff5f77ed8a4aac161167c0a8b18638e5ea5544",
  "setImpiantoDone": "ee75f2d0a699f9c141b47ce397ba9fe5b557fb20",
  "validateImpiantoCoordinates": "bed9d3049b3d39b85ac5feb9c20f247c713ead55",
  "canUseForceImpiantoDone": "25986435d4545ce1783ed78cf6825985b71ec437",
  "forceMarkDone": "3afcd87733a3188d3d8d353fd0c0a0ae71e8b354",
  "setImpiantoFattoSavingState": "e0f2e8a68a0cb849eb4dbca0e7955b3d980934ff",
  "handleImpiantoWhatsAppClick": "9e6e15701456e05d6e9012e78edb5caa10860a1b",
  "handleCompletedImpiantoWhatsAppClick": "56fa2be058f37297eb0bcedef8e951cdabc2b1a0",
  "forceMoveImpiantoToFatti": "2ad24b595eeaa1b94f4611ad31b1ed2dcfe641ab",
  "verifyImpiantoDoneBackground": "5446ef01e687f6d9f3c654d9574bdd6da20c4932",
  "runWhazzupPendingDoneSafetyCheck": "57eda7a5ecedce01b04fe0c6edbe7d904708f41a",
  "isImpiantoPersistedAsDone": "9af706f2b2b89337e74bf17d0a610fa01fe9328b",
  "getWhazzupSafetyState": "069bc7ee9095fa50e32ce5930a2b6e4fb510ba10",
  "markWhazzupSafetyPressed": "0d4c6dd226f9e14c7c81ed1c089b48b24933143e",
  "updateWhazzupSafetyAfterBackgroundCheck": "384ff3bbceca869a510dabfd752dbb321d57634d",
  "markImpiantoDoneRecoveryRequired": "6cf0620a5fc365fa2d83bdb1c836550a31c2bdea",
  "handleImpiantoDoneSaveFailure": "6f36840b13104a18de1e8a9a0c7f145e921fee69",
  "loadWhazzupPendingDoneEntries": "8860a04ed0966df1ff24351a5e4e8163d58795ed",
  "saveWhazzupPendingDoneEntries": "e76f65d6d23c41390f12a5204f5c32f0e8455e90",
  "upsertWhazzupPendingDoneEntry": "38256367a2a5e21d6c37c70a5e3438e9258c933d",
  "clearWhazzupPendingDoneEntry": "5015475ace8b97459aac37a5667b1efd0cc63a2c",
  "upsertPendingDoneAction": "aded657452dbc278d2e1a2d69596c003b32a799d",
  "removePendingDoneActionsForImpianto": "a84286cf16fb2ab409bce87cbc88512748365707",
  "markPendingActionStatus": "04eb5edf45d3258ba0fdd62f517baec1bd3b67c4",
  "updateImpiantoLocalState": "abaf285b7f765f66bdf41741576da6a8547cf836",
  "getImpiantoDocIds": "86a772688561b2ec65d7fb8be8e3d9fb441becbb",
  "safeOpenWhatsAppMessage": "b0bb9bab5458871395706a029f2d33d769747ebb",
  "openDeferredWhatsAppTargetWindow": "83422589acd483be8dea4f2382a905ff201bfecd",
  "getImpiantoWhatsAppTemplateSignature": "17f98fb77f6db75d400b63b0892d9ecc9541e9c4",
  "buildImpiantoWhatsAppTemplate": "f25916a8e01af18731b929078319721aa158a1b0",
  "prepareImpiantoWhatsAppTemplate": "4c5982042f05f577c4025bfc1833bfdadfbadf73",
  "refreshImpiantoWhatsAppTemplateCache": "c0d03c4a55e52aee10765d2f183f56ab54cd5262",
  "buildImpiantoWhatsAppPayload": "ea957d19f7bec801f01378adf85d4ccc5162438f",
  "shareWhazzupWithPhotos": "89d7954fed5e525ca1f789a556527243b5bb2f78",
  "openWhatsApp": "c6892a55b330c96bd00af727bcf77a5ff1527df0"
});
const LOCKED_FILES = Object.freeze({
  "fatto-button-immediate.js": "8f70bca841a0e75e5fc4750e31e2c9cff92e5167",
  "fatto-scroll-guard.js": "7ff46878e147d3826b354a4f89eebe0d1d636027",
  "native-android-runtime.js": "db0aa29bd26910398effd7ac9b0f15e26f8250bc",
  "android-whazzup-photo-order.js": "93ab728cade81ab633c08e34b0b60f5581057725",
  "whazzup-preload-cache.js": "d27ca8a36ae5a32ae9e90c9c572cb7d37c491fad",
  "android/app/src/main/java/it/vargacantieri/hera/whatsapp/HeraWhatsAppPlugin.java": "c1bf032d116e2a7adf32413e103872272ee848a8",
  "android/app/src/main/java/it/vargacantieri/hera/whatsapp/HeraWhazzupPhotosPlugin.java": "c0daf2d24dbbdad53f5408722568be88fef0050e",
  "scripts/check-done-button-logic.js": "4408d6ff4e2d49329819beb49609833e4140e49b",
  "scripts/check-done-button-behavior.js": "ab8bafb573effca872122bca2bac5544c5444f9a",
  "scripts/check-whatsapp-protection.js": "40b86c51e3d46ef0a340685f8c166b2b807c495d"
});

function fail(lines) {
  const message = Array.isArray(lines) ? lines.join("\n") : String(lines);
  console.error("\n❌ CASSAFORTE IRREVOCABILE FATTO/WHAZZUP BLOCCATA\n" + message + "\n");
  process.exit(1);
}

function gitBlobSha(value) {
  const buffer = Buffer.isBuffer(value) ? value : Buffer.from(String(value), "utf8");
  const header = Buffer.from(`blob ${buffer.length}\0`, "utf8");
  return crypto.createHash("sha1").update(Buffer.concat([header, buffer])).digest("hex");
}

function read(relativePath) {
  const absolutePath = path.join(root, relativePath);
  if (!fs.existsSync(absolutePath)) fail(`File protetto mancante: ${relativePath}`);
  return fs.readFileSync(absolutePath);
}

function extractFunction(source, name) {
  const signatures = [`async function ${name}(`, `function ${name}(`];
  let start = -1;
  for (const signature of signatures) {
    start = source.indexOf(signature);
    if (start >= 0) break;
  }
  if (start < 0) fail(`Funzione protetta mancante: ${name}`);
  const signatureEnd = source.indexOf(") {", start);
  const bodyStart = source.indexOf("{", signatureEnd);
  let depth = 0;
  for (let index = bodyStart; index < source.length; index += 1) {
    if (source[index] === "{") depth += 1;
    if (source[index] === "}") depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }
  fail(`Funzione protetta non chiusa: ${name}`);
}

const agents = read("AGENTS.md").toString("utf8");
if (!agents.includes(LOCK_MARKER)) {
  fail("Regola irrevocabile assente da AGENTS.md. Non è consentito rimuoverla o attenuarla.");
}

const appSource = read("app.js").toString("utf8");
for (const [name, expectedSha] of Object.entries(LOCKED_FUNCTIONS)) {
  const actualSha = gitBlobSha(extractFunction(appSource, name));
  if (actualSha !== expectedSha) {
    fail([
      `Funzione protetta modificata: ${name}`,
      `Impronta approvata: ${expectedSha}`,
      `Impronta trovata: ${actualSha}`,
      `Versione definitiva: commit ${LOCKED_AT_COMMIT}`,
      "La modifica è vietata anche con un futuro consenso dell'amministratore."
    ]);
  }
}

for (const [relativePath, expectedSha] of Object.entries(LOCKED_FILES)) {
  const actualSha = gitBlobSha(read(relativePath));
  if (actualSha !== expectedSha) {
    fail([
      `File protetto modificato: ${relativePath}`,
      `Impronta approvata: ${expectedSha}`,
      `Impronta trovata: ${actualSha}`,
      "Ripristinare la versione definitiva. Non aggiornare l'impronta."
    ]);
  }
}

console.log(`✅ Cassaforte irrevocabile FATTO/WHAZZUP attiva: ${Object.keys(LOCKED_FUNCTIONS).length} funzioni e ${Object.keys(LOCKED_FILES).length} file invariati.`);
console.log(`✅ Versione definitiva bloccata al commit ${LOCKED_AT_COMMIT}.`);

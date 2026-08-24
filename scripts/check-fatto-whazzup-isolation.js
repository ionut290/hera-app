#!/usr/bin/env node
"use strict";

const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const ISOLATION_MARKER = "PERIMETRO_ISOLATO_FATTO_WHAZZUP_V1";
const IMMUTABLE_LOCK_SHA = "982c5015da3590311fb52d4395e13387f3f2958e";
const CRITICAL_NAMES = Object.freeze([
  "isImpiantoDoneState",
  "combineImpiantiForView",
  "markImpiantoDone",
  "markImpiantoDoneVisualFallback",
  "retrySetImpiantoDone",
  "resetImpianto",
  "setImpiantoDone",
  "validateImpiantoCoordinates",
  "canUseForceImpiantoDone",
  "forceMarkDone",
  "setImpiantoFattoSavingState",
  "handleImpiantoWhatsAppClick",
  "handleCompletedImpiantoWhatsAppClick",
  "forceMoveImpiantoToFatti",
  "verifyImpiantoDoneBackground",
  "runWhazzupPendingDoneSafetyCheck",
  "isImpiantoPersistedAsDone",
  "getWhazzupSafetyState",
  "markWhazzupSafetyPressed",
  "updateWhazzupSafetyAfterBackgroundCheck",
  "markImpiantoDoneRecoveryRequired",
  "handleImpiantoDoneSaveFailure",
  "loadWhazzupPendingDoneEntries",
  "saveWhazzupPendingDoneEntries",
  "upsertWhazzupPendingDoneEntry",
  "clearWhazzupPendingDoneEntry",
  "upsertPendingDoneAction",
  "removePendingDoneActionsForImpianto",
  "markPendingActionStatus",
  "updateImpiantoLocalState",
  "getImpiantoDocIds",
  "safeOpenWhatsAppMessage",
  "openDeferredWhatsAppTargetWindow",
  "getImpiantoWhatsAppTemplateSignature",
  "buildImpiantoWhatsAppTemplate",
  "prepareImpiantoWhatsAppTemplate",
  "refreshImpiantoWhatsAppTemplateCache",
  "buildImpiantoWhatsAppPayload",
  "shareWhazzupWithPhotos",
  "openWhatsApp"
]);
const APPROVED_RUNTIME_FILES = new Set([
  "app.js",
  "fatto-button-immediate.js",
  "fatto-scroll-guard.js",
  "native-android-runtime.js",
  "whazzup-preload-cache.js",
  "android-whazzup-photo-order.js",
  "varga-branding.js"
]);
const ALL_CHANGE_WORKFLOWS = Object.freeze([
  ".github/workflows/check-done-button.yml",
  ".github/workflows/check-critical-flows.yml",
  ".github/workflows/check-e2e-smoke.yml"
]);

function fail(lines) {
  const message = Array.isArray(lines) ? lines.join("\n") : String(lines);
  console.error("\n❌ PERIMETRO ISOLATO FATTO/WHAZZUP BLOCCATO\n" + message + "\n");
  process.exit(1);
}

function read(relativePath) {
  const absolutePath = path.join(root, relativePath);
  if (!fs.existsSync(absolutePath)) fail(`File necessario mancante: ${relativePath}`);
  return fs.readFileSync(absolutePath);
}

function gitBlobSha(value) {
  const buffer = Buffer.isBuffer(value) ? value : Buffer.from(String(value), "utf8");
  const header = Buffer.from(`blob ${buffer.length}\0`, "utf8");
  return crypto.createHash("sha1").update(Buffer.concat([header, buffer])).digest("hex");
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function assertClassicOrderedScript(html, fileName) {
  const expression = new RegExp(`<script(?![^>]*\\b(?:async|defer|type=["']module["']))[^>]*src=["'][^"']*${escapeRegExp(fileName)}(?:\\?[^"']*)?["'][^>]*><\\/script>`, "g");
  const matches = Array.from(html.matchAll(expression));
  if (matches.length !== 1) {
    fail(`Lo script critico ${fileName} deve essere caricato una sola volta, in modo classico e sincrono.`);
  }
  return matches[0].index;
}

const agents = read("AGENTS.md").toString("utf8");
if (!agents.includes("BLOCCO_IRREVOCABILE_FATTO_WHAZZUP_V1") || !agents.includes(ISOLATION_MARKER)) {
  fail("Le regole permanenti della cassaforte o del perimetro isolato sono assenti.");
}

const immutableLock = read("scripts/check-fatto-whazzup-immutable.js");
if (gitBlobSha(immutableLock) !== IMMUTABLE_LOCK_SHA) {
  fail("La cassaforte primaria è stata modificata. Ripristinare l’impronta definitiva.");
}

const html = read("index.html").toString("utf8");
const nativeIndex = assertClassicOrderedScript(html, "native-android-runtime.js");
const appIndex = assertClassicOrderedScript(html, "app.js");
const immediateIndex = assertClassicOrderedScript(html, "fatto-button-immediate.js");
if (!(nativeIndex < appIndex && appIndex < immediateIndex)) {
  fail("Ordine critico alterato: native-android-runtime.js -> app.js -> fatto-button-immediate.js.");
}

const runtimeFiles = fs.readdirSync(root, { withFileTypes: true })
  .filter((entry) => entry.isFile() && entry.name.endsWith(".js"))
  .map((entry) => entry.name)
  .sort();

for (const fileName of runtimeFiles) {
  if (APPROVED_RUNTIME_FILES.has(fileName)) continue;
  const source = read(fileName).toString("utf8");
  for (const name of CRITICAL_NAMES) {
    const escaped = escapeRegExp(name);
    const forbiddenPatterns = [
      new RegExp(`\\b(?:async\\s+)?function\\s+${escaped}\\s*\\(`),
      new RegExp(`(?:window|globalThis)\\s*\\.\\s*${escaped}\\s*=`),
      new RegExp(`(?:window|globalThis)\\s*\\[\\s*["']${escaped}["']\\s*\\]\\s*=`),
      new RegExp(`Object\\.defineProperty\\s*\\(\\s*(?:window|globalThis)\\s*,\\s*["']${escaped}["']`),
      new RegExp(`delete\\s+(?:window|globalThis)\\s*\\.\\s*${escaped}\\b`)
    ];
    if (forbiddenPatterns.some((pattern) => pattern.test(source))) {
      fail([
        `Interferenza esterna rilevata in ${fileName}`,
        `La funzione protetta ${name} non può essere ridefinita, sostituita, cancellata o intercettata.`
      ]);
    }
  }
}

for (const fileName of runtimeFiles) {
  const source = read(fileName).toString("utf8");
  if (fileName !== "app.js" && (source.includes('window.open("about:blank"') || source.includes("Preparazione messaggio in corso"))) {
    fail(`Pagina intermedia Whazzup reintrodotta da ${fileName}.`);
  }
}

for (const workflowPath of ALL_CHANGE_WORKFLOWS) {
  const workflow = read(workflowPath).toString("utf8");
  if (/^\s+paths(?:-ignore)?:\s*$/m.test(workflow)) {
    fail(`Il workflow ${workflowPath} deve eseguire i controlli su ogni modifica, senza filtri paths.`);
  }
  if (!workflow.includes("check-fatto-whazzup-isolation.js")) {
    fail(`Il workflow ${workflowPath} non esegue il perimetro isolato.`);
  }
}

const packageJson = JSON.parse(read("package.json").toString("utf8"));
if (packageJson.scripts?.["check:fatto-isolation"] !== "node scripts/check-fatto-whazzup-isolation.js") {
  fail("Script npm check:fatto-isolation mancante o alterato.");
}
if (!String(packageJson.scripts?.["check:fatto-critical"] || "").startsWith("npm run check:fatto-whazzup-immutable && npm run check:fatto-isolation &&")) {
  fail("Il controllo critico non esegue prima cassaforte e isolamento.");
}

console.log(`✅ Perimetro FATTO/WHAZZUP isolato da ${runtimeFiles.length} file runtime esterni.`);
console.log("✅ Ogni modifica futura attiva cassaforte, controlli critici e browser E2E.");
console.log("✅ Ordine di caricamento e assenza di sostituzioni esterne verificati.");

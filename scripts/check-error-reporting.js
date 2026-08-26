"use strict";

const fs = require("node:fs");
const assert = require("node:assert/strict");

const client = fs.readFileSync("client-error-reporter.js", "utf8");
const monitor = fs.readFileSync("app-error-monitor.js", "utf8");
const adminCenter = fs.readFileSync("admin-error-center.js", "utf8");
const emailBackend = fs.readFileSync("functions/error-reporting.js", "utf8");
const centerBackend = fs.readFileSync("functions/error-center.js", "utf8");
const main = fs.readFileSync("functions/main.js", "utf8");
const loader = fs.readFileSync("loading-humor.js", "utf8");
const serviceWorker = fs.readFileSync("sw.js", "utf8");
const headers = fs.readFileSync("_headers", "utf8");
const deployWorkflow = fs.readFileSync(".github/workflows/deploy-firebase-functions.yml", "utf8");

// Reporter email preesistente: deve restare isolato e silenzioso.
assert.match(client, /addEventListener\("error"/);
assert.match(client, /unhandledrejection/);
assert.match(client, /hera_client_error_queue_v1/);
assert.match(client, /hera_client_error_dedupe_v1/);
assert.match(client, /httpsCallable\(FUNCTION_NAME\)/);
assert.match(client, /europe-west1/);
assert.match(client, /monthlyLimited/);
assert.doesNotMatch(client, /showSentToast/);
assert.doesNotMatch(client, /Diagnosi inviata automaticamente/);
assert.doesNotMatch(client, /\.firestore\s*\(/);
assert.doesNotMatch(client, /onSnapshot\s*\(/);
assert.doesNotMatch(client, /setInterval\s*\(/);

// Monitor globale: errori, lentezza, blocchi, tocchi ripetuti e coda offline.
assert.match(monitor, /recordClientErrorGroup/);
assert.match(monitor, /hera_error_center_queue_v1/);
assert.match(monitor, /hera_error_center_dedupe_v1/);
assert.match(monitor, /addEventListener\("error"/);
assert.match(monitor, /unhandledrejection/);
assert.match(monitor, /repeated-tap/);
assert.match(monitor, /slow-interaction/);
assert.match(monitor, /ui-freeze/);
assert.match(monitor, /PerformanceObserver/);
assert.match(monitor, /reportManual/);
assert.match(monitor, /handled-console-error/);
assert.match(monitor, /getHealth/);
assert.match(monitor, /lastSuccessfulSendAt/);
assert.match(monitor, /SENSITIVE_KEY/);
assert.match(monitor, /CHIAVE_API_RIMOSSA/);
assert.doesNotMatch(monitor, /setInterval\s*\(/);
assert.doesNotMatch(monitor, /target\.value|actionable\.value|event\.target\.value/);
assert.doesNotMatch(monitor, /geolocation|getCurrentPosition|latitude|longitude/);

// Centro amministratore: solo callable mirate, nessun listener Firestore autonomo.
assert.match(adminCenter, /open-admin-error-center-btn/);
assert.match(adminCenter, /open-app-bug-report-btn/);
assert.match(adminCenter, /getErrorCenterSummary/);
assert.match(adminCenter, /getErrorCenterDashboard/);
assert.match(adminCenter, /markErrorCenterSeen/);
assert.match(adminCenter, /updateErrorCenterStatus/);
assert.match(adminCenter, /ERROR_CENTER/);
assert.match(adminCenter, /HeraAppErrorMonitor/);
assert.match(adminCenter, /countsVerified/);
assert.match(adminCenter, /Dati non verificati/);
assert.match(adminCenter, /Gli zeri non vengono mostrati/);
assert.doesNotMatch(adminCenter, /onSnapshot\s*\(/);
assert.doesNotMatch(adminCenter, /setInterval\s*\(/);
assert.doesNotMatch(adminCenter, /\.collection\s*\(/);

// Backend: aggregazione e autorizzazione. Il Centro Errori deve restare silenzioso:
// registra e mostra gli errori nell'app, senza creare notifiche o push amministrativi.
assert.match(centerBackend, /appErrorGroups/);
assert.match(centerBackend, /systemCounters/);
assert.match(centerBackend, /runTransaction/);
assert.match(centerBackend, /recordClientErrorGroup/);
assert.match(centerBackend, /getErrorCenterSummary/);
assert.match(centerBackend, /getErrorCenterDashboard/);
assert.match(centerBackend, /PUBLIC_CALLABLE_OPTIONS\s*=\s*Object\.freeze\(\{\s*invoker:\s*"public"\s*\}\)/);
assert.equal(
  (centerBackend.match(/runWith\(PUBLIC_CALLABLE_OPTIONS\)\.https\.onCall/g) || []).length,
  5,
  "all five Error Center callable functions must explicitly allow public transport access"
);
assert.match(centerBackend, /markErrorCenterSeen/);
assert.match(centerBackend, /updateErrorCenterStatus/);
assert.match(centerBackend, /countSource:\s*"firestore-aggregate"/);
assert.match(centerBackend, /countsVerified:\s*true/);
assert.match(centerBackend, /context\.auth/);
assert.doesNotMatch(centerBackend, /collection\(["']notifications["']\)/);
assert.doesNotMatch(centerBackend, /sendEachForMulticast/);
assert.doesNotMatch(centerBackend, /scopeType:\s*["']ADMIN["']/);
assert.doesNotMatch(centerBackend, /actionType:\s*["']ERROR_CENTER["']/);
assert.doesNotMatch(centerBackend, /onSnapshot\s*\(/);
assert.doesNotMatch(centerBackend, /setInterval\s*\(/);

// Il deploy deve applicare e rileggere la policy IAM senza rimuovere gli altri binding.
assert.match(deployWorkflow, /Grant and verify Error Center callable transport/);
assert.ok(deployWorkflow.includes("roles/cloudfunctions.invoker"));
assert.match(deployWorkflow, /allUsers/);
assert.match(deployWorkflow, /:getIamPolicy/);
assert.match(deployWorkflow, /:setIamPolicy/);
for (const callableName of [
  "recordClientErrorGroup",
  "getErrorCenterSummary",
  "getErrorCenterDashboard",
  "markErrorCenterSeen",
  "updateErrorCenterStatus"
]) {
  assert.ok(
    deployWorkflow.includes(`"${callableName}"`),
    `${callableName} deve ricevere il binding invoker durante il deploy`
  );
}

// Sistema email preesistente protetto.
assert.match(emailBackend, /defineSecret\("RESEND_API_KEY"\)/);
assert.match(emailBackend, /defineSecret\("ERROR_REPORT_FROM"\)/);
assert.match(emailBackend, /MONTHLY_EMAIL_LIMIT = 2500/);
assert.match(emailBackend, /api\.resend\.com\/emails/);

assert.match(main, /require\("\.\/error-reporting"\)/);
assert.match(main, /require\("\.\/error-center"\)/);
assert.match(main, /errorCenterFunctions/);
assert.match(loader, /client-error-reporter\.js\?v=20260816a/);
assert.match(loader, /app-error-monitor\.js\?v=20260824b/);
assert.match(loader, /admin-error-center\.js\?v=20260824b/);
assert.match(loader, /admin-error-center\.css\?v=20260824b/);
assert.match(serviceWorker, /varga-cantieri-shell-v158/);
for (const path of ["/loading-humor.js", "/client-error-reporter.js", "/app-error-monitor.js", "/admin-error-center.js", "/admin-error-center.css"]) {
  assert.ok(serviceWorker.includes(`"${path}"`), `${path} deve essere network-first`);
}
for (const path of ["/client-error-reporter.js", "/app-error-monitor.js", "/admin-error-center.js", "/admin-error-center.css"]) {
  assert.ok(headers.includes(path), `${path} deve avere intestazione no-cache`);
}

const protectedNames = [
  "setImpiantoDone",
  "markImpiantoDone",
  "openWhatsApp",
  "openWhazzup",
  "buildWhatsApp",
  "buildWhazzup"
];
for (const name of protectedNames) {
  for (const [label, source] of [["monitor", monitor], ["centro amministratore", adminCenter], ["backend centro", centerBackend]]) {
    assert.equal(source.includes(name), false, `${label} non deve riferirsi a ${name}`);
  }
}

console.log("Centro errori: controlli statici, privacy, modalità silenziosa e isolamento superati.");

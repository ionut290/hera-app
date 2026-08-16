"use strict";

const fs = require("node:fs");
const assert = require("node:assert/strict");

const client = fs.readFileSync("client-error-reporter.js", "utf8");
const backend = fs.readFileSync("functions/error-reporting.js", "utf8");
const main = fs.readFileSync("functions/main.js", "utf8");
const loader = fs.readFileSync("loading-humor.js", "utf8");

assert.match(client, /addEventListener\("error"/);
assert.match(client, /unhandledrejection/);
assert.match(client, /hera_client_error_queue_v1/);
assert.match(client, /hera_client_error_dedupe_v1/);
assert.match(client, /httpsCallable\(FUNCTION_NAME\)/);
assert.match(client, /europe-west1/);
assert.doesNotMatch(client, /\.firestore\s*\(/);
assert.doesNotMatch(client, /onSnapshot\s*\(/);
assert.doesNotMatch(client, /setInterval\s*\(/);

assert.match(backend, /defineSecret\("RESEND_API_KEY"\)/);
assert.match(backend, /defineSecret\("ERROR_REPORT_FROM"\)/);
assert.match(backend, /runWith\(\{ secrets: \[RESEND_API_KEY, ERROR_REPORT_FROM\] \}\)/);
assert.match(backend, /context\.auth/);
assert.match(backend, /Idempotency-Key/);
assert.match(backend, /api\.resend\.com\/emails/);
assert.doesNotMatch(backend, /\.firestore\s*\(/);
assert.doesNotMatch(backend, /collection\s*\(/);
assert.doesNotMatch(backend, /onSnapshot\s*\(/);
assert.doesNotMatch(backend, /setInterval\s*\(/);

assert.match(main, /require\("\.\/error-reporting"\)/);
assert.match(main, /errorReportingFunctions/);
assert.match(loader, /client-error-reporter\.js\?v=20260816a/);

const protectedNames = [
  "setImpiantoDone",
  "markImpiantoDone",
  "openWhatsApp",
  "openWhazzup",
  "buildWhatsApp",
  "buildWhazzup"
];
for (const name of protectedNames) {
  assert.equal(client.includes(name), false, `Il client diagnostico non deve riferirsi a ${name}`);
  assert.equal(backend.includes(name), false, `Il backend diagnostico non deve riferirsi a ${name}`);
}

console.log("Diagnostica errori: controlli statici superati.");

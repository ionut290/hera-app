"use strict";
// Tests run exclusively against a local Firebase Storage emulator.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const { initializeTestEnvironment, assertFails, assertSucceeds } = require("@firebase/rules-unit-testing");
const { ref, uploadString, getMetadata } = require("firebase/storage");

async function main() {
  const hostPort = process.env.FIREBASE_STORAGE_EMULATOR_HOST || "";
  const [host, portString] = hostPort.split(":");
  if (!["127.0.0.1", "localhost"].includes(host)) {
    throw new Error("Storage emulator must run locally; production is not permitted.");
  }
  const context = await initializeTestEnvironment({
    projectId: "demo-review-isolation",
    storage: {
      host, port: Number(portString),
      rules: fs.readFileSync("storage.rules", "utf8")
    }
  });
  try {
    const reviewer = context.authenticatedContext("reviewerUser", {
      email: "reviewer@example.test", email_verified: true, apple_review: true
    }).storage();
    const regular = context.authenticatedContext("normalUser", {
      email: "normal@example.test", email_verified: true
    }).storage();

    const reviewOwnedPath = "documents/reviewerUser/demoDoc/sample.txt";
    const reviewLegacyPath = "lavori-occasionali/demoPlant/reviewerUser/demoDoc/sample.pdf";
    const regularOwnedPath = "documents/normalUser/demoDoc/sample.txt";

    // Existing normal functionality must still be allowed for an owner.
    await assertSucceeds(uploadString(ref(regular, regularOwnedPath), "normal user's sample"));
    await assertSucceeds(getMetadata(ref(regular, regularOwnedPath)));

    // App Review is categorically forbidden from uploading or reading files.
    await assertFails(uploadString(ref(reviewer, reviewOwnedPath), "forbidden"));
    await assertFails(uploadString(ref(reviewer, reviewLegacyPath), "forbidden"));
    await assertFails(getMetadata(ref(reviewer, regularOwnedPath)));
    await assertFails(getMetadata(ref(reviewer, reviewOwnedPath)));
    await assertFails(getMetadata(ref(reviewer, reviewLegacyPath)));

    console.log("Apple reviewer Storage access is denied; normal-owner upload/read still works.");
  } finally {
    await context.cleanup();
  }
}

main().catch((error) => { console.error(error); process.exitCode = 1; });

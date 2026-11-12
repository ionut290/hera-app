"use strict";

// Run only in a local Firestore emulator; never against a Firebase project.
// Example: npx firebase emulators:exec --only firestore --project demo-review-isolation //   "node scripts/test-apple-review-firestore-rules.cjs"
const fs = require("node:fs");
const assert = require("node:assert/strict");
const { initializeTestEnvironment, assertFails, assertSucceeds } =
  require("@firebase/rules-unit-testing");

async function run() {
  if (!process.env.FIRESTORE_EMULATOR_HOST) {
    throw new Error("FIRESTORE_EMULATOR_HOST is missing. Never run against production.");
  }
  const [host, port] = process.env.FIRESTORE_EMULATOR_HOST.split(":");
  assert.ok(host === "127.0.0.1" || host === "localhost", "Only a local emulator is permitted.");
  const testEnv = await initializeTestEnvironment({
    projectId: "demo-review-isolation",
    firestore: {
      host,
      port: Number(port),
      rules: fs.readFileSync("firestore.rules", "utf8")
    }
  });

  try {
    await testEnv.withSecurityRulesDisabled(async (context) => {
      const db = context.firestore();
      await db.doc("platformUsers/normalUser").set({
        uid: "normalUser", email: "normal@example.test",
        statoAccount: "attivo", accountStatus: "attivo", role: "user"
      });
      await db.doc("platformUsers/reviewerUser").set({
        uid: "reviewerUser", email: "reviewer@example.test",
        statoAccount: "attivo", accountStatus: "attivo", role: "user"
      });
      await db.doc("commesse/syntheticJob").set({ name: "Synthetic fixture" });
      await db.doc("commesse/syntheticJob/impianti/testPlant").set({ name: "Plant" });
      await db.doc("personale/staff").set({ displayName: "Private staff" });
      await db.doc("utenti/staff").set({ displayName: "Private user" });
      await db.doc("oreReports/one").set({ date: "2026-11-12" });
      await db.doc("sharedStaticViews/currentTeam").set({ demo: false });
      await db.doc("appleReviewDemo/onlySynthetic").set({title: "Synthetic exercise", demo: true});
    });

    const normal = testEnv.authenticatedContext("normalUser", {
      email: "normal@example.test", email_verified: true
    }).firestore();
    const review = testEnv.authenticatedContext("reviewerUser", {
      email: "reviewer@example.test", email_verified: true, apple_review: true
    }).firestore();
    // Even an accidentally whitelisted admin address must never bypass the claim.
    const reviewAdminEmail = testEnv.authenticatedContext("reviewerUser", {
      email: "ionut29019@gmail.com", email_verified: true, apple_review: true
    }).firestore();

    await assertSucceeds(review.doc("appleReviewDemo/onlySynthetic").get());
    await assertSucceeds(review.collection("appleReviewDemo").get());
    await assertFails(review.doc("appleReviewDemo/onlySynthetic").set({ demo: false }));
    await assertFails(review.doc("appleReviewDemo/onlySynthetic").update({ demo: false }));
    await assertFails(review.doc("appleReviewDemo/onlySynthetic").delete());
    await assertFails(normal.doc("appleReviewDemo/onlySynthetic").get());

    await assertSucceeds(normal.doc("commesse/syntheticJob").get());
    await assertSucceeds(normal.doc("sharedStaticViews/currentTeam").get());
    await assertSucceeds(normal.doc("commesse/syntheticJob/impianti/testPlant").get());

    for (const docPath of [
      "commesse/syntheticJob",
      "commesse/syntheticJob/impianti/testPlant",
      "personale/staff",
      "utenti/staff",
      "oreReports/one",
      "sharedStaticViews/currentTeam",
      "platformUsers/reviewerUser"
    ]) {
      await assertFails(review.doc(docPath).get());
      await assertFails(review.doc(docPath).set({ test: "attempt" }));
      await assertFails(review.doc(docPath).update({ test: "attempt" }));
      await assertFails(review.doc(docPath).delete());
    }

    await assertFails(reviewAdminEmail.doc("appConfig/adminUsers").get());
    await assertFails(reviewAdminEmail.doc("appConfig/adminUsers").set({ emails: [] }));
    await assertFails(review.collection("commesse").get());
    await assertFails(review.collection("personale").get());
    console.log("Apple reviewer can only read synthetic demo data; live Firestore read/write denied and normal access unchanged.");
  } finally {
    await testEnv.cleanup();
  }
}
run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

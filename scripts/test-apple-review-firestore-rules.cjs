"use strict";

// Run only in a local Firestore emulator; never against a Firebase project.
// Example: npx firebase emulators:exec --only firestore --project demo-review-isolation //   "node scripts/test-apple-review-firestore-rules.cjs"
const fs = require("node:fs");
const { doc, collection, getDoc, getDocs, setDoc, updateDoc, deleteDoc } = require("firebase/firestore");
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
      await setDoc(doc(db, "platformUsers/normalUser"), {
        uid: "normalUser", email: "normal@example.test",
        statoAccount: "attivo", accountStatus: "attivo", role: "user"
      });
      await setDoc(doc(db, "platformUsers/reviewerUser"), {
        uid: "reviewerUser", email: "reviewer@example.test",
        statoAccount: "attivo", accountStatus: "attivo", role: "user"
      });
      await setDoc(doc(db, "commesse/syntheticJob"), { name: "Synthetic fixture" });
      await setDoc(doc(db, "commesse/syntheticJob/impianti/testPlant"), { name: "Plant" });
      await setDoc(doc(db, "personale/staff"), { displayName: "Private staff" });
      await setDoc(doc(db, "utenti/staff"), { displayName: "Private user" });
      await setDoc(doc(db, "oreReports/one"), { date: "2026-11-12" });
      await setDoc(doc(db, "sharedStaticViews/currentTeam"), { demo: false });
      await setDoc(doc(db, "appleReviewDemo/onlySynthetic"), {title: "Synthetic exercise", demo: true});
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

    await assertSucceeds(getDoc(doc(review, "appleReviewDemo/onlySynthetic")));
    await assertSucceeds(getDocs(collection(review, "appleReviewDemo")));
    await assertFails(setDoc(doc(review, "appleReviewDemo/onlySynthetic"), { demo: false }));
    await assertFails(updateDoc(doc(review, "appleReviewDemo/onlySynthetic"), { demo: false }));
    await assertFails(deleteDoc(doc(review, "appleReviewDemo/onlySynthetic")));
    await assertFails(getDoc(doc(normal, "appleReviewDemo/onlySynthetic")));

    await assertSucceeds(getDoc(doc(normal, "commesse/syntheticJob")));
    await assertSucceeds(getDoc(doc(normal, "sharedStaticViews/currentTeam")));
    await assertSucceeds(getDoc(doc(normal, "commesse/syntheticJob/impianti/testPlant")));

    for (const docPath of [
      "commesse/syntheticJob",
      "commesse/syntheticJob/impianti/testPlant",
      "personale/staff",
      "utenti/staff",
      "oreReports/one",
      "sharedStaticViews/currentTeam",
      "platformUsers/reviewerUser"
    ]) {
      await assertFails(getDoc(doc(review, docPath)));
      await assertFails(setDoc(doc(review, docPath), { test: "attempt" }));
      await assertFails(updateDoc(doc(review, docPath), { test: "attempt" }));
      await assertFails(deleteDoc(doc(review, docPath)));
    }

    await assertFails(getDoc(doc(reviewAdminEmail, "appConfig/adminUsers")));
    await assertFails(setDoc(doc(reviewAdminEmail, "appConfig/adminUsers"), { emails: [] }));
    await assertFails(getDocs(collection(review, "commesse")));
    await assertFails(getDocs(collection(review, "personale")));
    console.log("Apple reviewer can only read synthetic demo data; live Firestore read/write denied and normal access unchanged.");
  } finally {
    await testEnv.cleanup();
  }
}
run().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

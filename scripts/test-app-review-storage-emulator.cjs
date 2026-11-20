"use strict";

// Never touch production: all test identities, Firestore metadata and files are
// fictitious, and both emulators MUST run locally on a demo- project.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const {initializeTestEnvironment, assertSucceeds, assertFails} =
  require("@firebase/rules-unit-testing");
const {doc, setDoc} = require("firebase/firestore");
const {ref, uploadBytes, getBytes, deleteObject} = require("firebase/storage");

const PROJECT = "demo-apple-storage-isolation";
const BUCKET = "gs://" + PROJECT + ".appspot.com";
const bytes = new Uint8Array([37, 80, 68, 70, 45, 49, 46, 55, 10]);
const checkLocal = (name) => {
  const value = process.env[name] || "";
  assert.match(value, /^(localhost|127\.0\.0\.1):[0-9]+$/,
    name + " must point to a loopback emulator");
  const [host, port] = value.split(":");
  return {host, port: Number(port)};
};

async function main() {
  assert.equal(process.env.GOOGLE_CLOUD_PROJECT, PROJECT,
    "Refuse production / non-demo Firebase projects");
  const storageHost = checkLocal("FIREBASE_STORAGE_EMULATOR_HOST");
  const firestoreHost = checkLocal("FIRESTORE_EMULATOR_HOST");
  const env = await initializeTestEnvironment({
    projectId: PROJECT,
    firestore: {...firestoreHost, rules: fs.readFileSync("firestore.rules", "utf8")},
    storage: {...storageHost, rules: fs.readFileSync("storage.rules", "utf8")}
  });
  try {
    await env.withSecurityRulesDisabled(async (context) => {
      const db = context.firestore();
      for (const [id, visibility, sharedToAll, sharedUserIds] of [
        ["testPersonal", "personal", false, []],
        ["testGlobal", "global", true, []],
        ["testShared", "shared", false, ["reviewColleague"]]
      ]) {
        await setDoc(doc(db, "documents/" + id), {
          id, ownerUserId: "reviewOwner", visibility, sharedToAll,
          sharedUserIds, synthetic: true
        });
      }
    });

    const owner = env.authenticatedContext("reviewOwner", {
      email: "owner@example.test", email_verified: true
    }).storage(BUCKET);
    const colleague = env.authenticatedContext("reviewColleague", {
      email: "colleague@example.test", email_verified: true
    }).storage(BUCKET);
    const unrelated = env.authenticatedContext("reviewUnrelated", {
      email: "unrelated@example.test", email_verified: true
    }).storage(BUCKET);
    const reviewer = env.authenticatedContext("reviewReviewer", {
      email: "reviewer@example.test", email_verified: true,
      apple_review: true, appReviewDemo: true
    }).storage(BUCKET);
    const fakeAdminReviewer = env.authenticatedContext("reviewReviewer", {
      email: "ionut29019@gmail.com", email_verified: true,
      apple_review: true, appReviewDemo: true, admin: true
    }).storage(BUCKET);
    const guest = env.unauthenticatedContext().storage(BUCKET);

    const pathFor = (id) => "documents/reviewOwner/" + id + "/fixture.pdf";
    const personal = pathFor("testPersonal");
    const global = pathFor("testGlobal");
    const shared = pathFor("testShared");
    const legacy = "lavori-occasionali/fakePlant/reviewOwner/fakeJob/fixture.pdf";
    const reference = (storage, path) => ref(storage, path);
    const uploadPdf = (storage, path) =>
      uploadBytes(reference(storage, path), bytes, {contentType: "application/pdf"});

    // Normal document upload is owner-only. A reviewer can never read or write
    // the normal document bucket, even if they spoof an admin email/claim.
    await assertSucceeds(uploadPdf(owner, personal));
    await assertSucceeds(uploadPdf(owner, global));
    await assertSucceeds(uploadPdf(owner, shared));
    await assertSucceeds(getBytes(reference(owner, personal)));
    await assertFails(getBytes(reference(colleague, personal)));
    await assertFails(getBytes(reference(unrelated, personal)));
    await assertFails(getBytes(reference(guest, personal)));
    await assertFails(uploadPdf(colleague, "documents/reviewOwner/other/file.pdf"));
    for (const storage of [reviewer, fakeAdminReviewer]) {
      await assertFails(getBytes(reference(storage, personal)));
      await assertFails(getBytes(reference(storage, global)));
      await assertFails(getBytes(reference(storage, shared)));
      await assertFails(uploadPdf(storage, "documents/reviewReviewer/test/file.pdf"));
      await assertFails(deleteObject(reference(storage, personal)));
    }

    // Shared/global access depends on the metadata in Firestore; it is still
    // forbidden for claimed Apple reviewers regardless of data visibility.
    await assertSucceeds(getBytes(reference(colleague, global)));
    await assertSucceeds(getBytes(reference(unrelated, global)));
    await assertSucceeds(getBytes(reference(colleague, shared)));
    await assertFails(getBytes(reference(unrelated, shared)));

    // Legacy PDF attachments have a stricter MIME policy and owner-only reads.
    await assertSucceeds(uploadPdf(owner, legacy));
    await assertSucceeds(getBytes(reference(owner, legacy)));
    await assertFails(getBytes(reference(colleague, legacy)));
    await assertFails(getBytes(reference(reviewer, legacy)));
    await assertFails(uploadBytes(
      reference(owner, "lavori-occasionali/fakePlant/reviewOwner/fakeJob/bad.txt"),
      bytes, {contentType: "text/plain"}
    ));

    // Delete only our fictitious files; no hidden Firebase credentials.
    for (const path of [personal, global, shared, legacy]) {
      await assertSucceeds(deleteObject(reference(owner, path)));
    }
    console.log("PASS: local Storage emulator owner/shared/global/legacy rules");
    console.log("PASS: reviewer, spoofed-admin reviewer, guest and unrelated access isolation");
    console.log("NOT PROVEN: real iPhone PDF upload or fully functional Apple review login.");
  } finally {
    await env.cleanup();
  }
}

main().catch(error => {
  console.error("Local Storage isolation verification FAILED:", error?.message || "unknown");
  process.exitCode = 1;
});

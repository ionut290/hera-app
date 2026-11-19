"use strict";

// Read-only production gate. Never creates, deletes or modifies Firestore data.
// In CI use --self-test (fixtures only). In protected manual release, authenticate
// with a short-lived service account file to verify both index states are READY.
const assert = require("node:assert/strict");
const fs = require("node:fs");

const PROJECT = "hera-app-6cd2b";
const REQUIRED = ["recipientId", "senderId"];

function hasReadyIndex(indexes, field) {
  return indexes.some((index) => {
    const fields = Array.isArray(index.fields) ? index.fields : [];
    const important = fields.filter(f => f.fieldPath !== "__name__");
    return index.state === "READY" &&
      index.queryScope === "COLLECTION" &&
      important.length === 2 &&
      important[0].fieldPath === field &&
      important[0].order === "ASCENDING" &&
      important[1].fieldPath === "createdAt" &&
      important[1].order === "DESCENDING";
  });
}

function verify(indexes) {
  for (const field of REQUIRED) {
    assert.ok(hasReadyIndex(indexes, field),
      "Missing READY composite chatMessages index: " + field + " ASC, createdAt DESC");
  }
}

async function main() {
  if (process.argv.includes("--self-test")) {
    const ready = (field) => ({
      state: "READY", queryScope: "COLLECTION",
      fields: [{fieldPath: field, order: "ASCENDING"},
        {fieldPath: "createdAt", order: "DESCENDING"},
        {fieldPath: "__name__", order: "DESCENDING"}]
    });
    verify([ready("recipientId"), ready("senderId")]);
    assert.throws(() => verify([ready("recipientId")]), /senderId/);
    assert.throws(() => verify([
      ready("recipientId"), {...ready("senderId"), state: "CREATING"}
    ]), /senderId/);
    assert.throws(() => verify([
      ready("recipientId"), {...ready("senderId"), queryScope: "COLLECTION_GROUP"}
    ]), /senderId/);
    console.log("PASS: Firestore chat composite indexes READY gate (synthetic fixtures).");
    return;
  }

  const keyFile = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  if (!keyFile) throw Error("Read-only protected release requires a service account.");
  const json = JSON.parse(fs.readFileSync(keyFile, "utf8"));
  if (json.project_id !== PROJECT) throw Error("Wrong Firebase service-account project.");
  // Only import credentials-backed client after fail-closed project validation.
  const {GoogleAuth} = require("google-auth-library");
  const auth = new GoogleAuth({
    keyFile, scopes: ["https://www.googleapis.com/auth/datastore"]
  });
  const client = await auth.getClient();
  let next = "";
  const indexes = [];
  do {
    const url = "https://firestore.googleapis.com/v1/projects/" + PROJECT +
      "/databases/(default)/collectionGroups/chatMessages/indexes?pageSize=100" +
      (next ? "&pageToken=" + encodeURIComponent(next) : "");
    const response = await client.request({method: "GET", url});
    indexes.push(...(response.data?.indexes || []));
    next = String(response.data?.nextPageToken || "");
    if (indexes.length > 1000) throw Error("Unexpected chat index pagination.");
  } while (next);
  verify(indexes);
  console.log("PASS: both private chat indexes are READY in expected Firebase project.");
}
main().catch((error) => {
  console.error("Private chat index safety gate failed:", error.message);
  process.exitCode = 1;
});

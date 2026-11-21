"use strict";

/**
 * Static review inventory ONLY. No Firebase SDK, credentials, network access,
 * record contents, deletion or legal conclusions. An automated test of source
 * coverage must never be misrepresented as completed GDPR/privacy sign-off.
 */
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const root = path.resolve(__dirname, "..");
const source = name => fs.readFileSync(path.join(root, name), "utf8");
const deletion = source("functions/account-deletion.js");
const firestoreRules = source("firestore.rules");
const privacy = source("privacy.html");
const retention = source("scripts/audit-retention-readiness.cjs");

/**
 * Explicit review inventory of the collection names observed in the existing
 * account-deletion implementation. Descriptions are CODE BEHAVIOR, not legal
 * grounds for processing or statements that erasure is complete.
 *
 * Each collection must stay in this manifest when the deletion code changes;
 * future collections still require manual schema review and may be missing.
 */
const deletionInventory = Object.freeze([
  ["documents", "delete-personal-only; retain-shared"],
  ["documentReminderAcknowledgements", "delete-own-query"],
  ["alertReadConfirmations", "delete-own-query"],
  ["userAlertAcknowledgements", "delete-own-query"],
  ["userAlerts", "delete-own-query"],
  ["passwordRecoveryChallenges", "delete-own-query"],
  ["personale", "unlink-auth-derived-fields-only"],
  ["notifications", "unlink-identity-and-personal-delivery-state"],
  ["chatMessages", "unlink-author-and-recipient-id"],
  ["appNotifications", "unlink-author-identity"],
  ["notificationEmailQueue", "minimize-exact-user-address"],
  ["activityLogs", "minimize-exact-user-identity"],
  ["privateDocuments", "delete-own-document"],
  ["operatorPositions", "delete-own-document"],
  ["userOperatorLinks", "delete-own-document"],
  ["platformUsers", "delete-own-document"]
]);

// Collections with business information are intentionally NOT removed by the
// self-service callable. Their content may still contain personal information.
const outsideDeletion = Object.freeze([
  ["commesse", "business-retention-needs-review"],
  ["globalCommesse", "business-retention-needs-review"],
  ["oreReports", "working-time-retention-needs-review"],
  ["calendarEvents", "events-can-contain-personal-details"]
]);

const unresolved = Object.freeze([
  "Controller/processor and lawful retention basis not independently signed off",
  "Firestore document/field inventory beyond the listed collections not exhaustive",
  "Operational files and thumbnails retained on shared references require review",
  "Legacy aliases, free-text identity, chat and historical exports not fully verified",
  "Google Drive copies, backups, off-platform email and third-party recipients unverified",
  "Ordinary-account deletion is deployed for authenticated users; verify legacy data, retained business records, and disposable account end-to-end",
  "Apple reauthorization/revocation and deletion not fully verified on physical iPhone"
]);

function audit() {
  // Track the actual released feature without treating this static inventory
  // as legal sign-off or proof of complete erasure of off-platform copies.
  assert.ok(deletion.includes("const DEMO_ONLY_RELEASE = false;"),
    "General deletion release mode changed; re-review inventory");
  assert.ok(deletion.includes("const uid = auth && auth.uid;"),
    "Account deletion must be scoped to the authenticated UID");
  assert.ok(deletion.includes("const MAX_AUTH_AGE_SECONDS = 600;"),
    "Recent authentication gate changed; re-review inventory");
  assert.ok(deletion.includes("ownAuthRecord.uid !== uid || ownAuthRecord.disabled === true"),
    "Live Firebase Auth record verification changed; re-review inventory");
  assert.ok(deletion.includes('keys.every(key => key === "confirmation" || key === "mode")'),
    "Deletion request payload allow-list changed; re-review inventory");
  assert.ok(deletion.includes("await admin.auth().deleteUser(uid);"),
    "Auth deletion last-step assumption changed; re-review inventory");
  assert.ok(deletion.includes('if (request.data?.mode === "preflight")'),
    "Read-only preflight contract changed; re-review inventory");

  const found = new Set();
  for (const [name] of deletionInventory) {
    assert.ok(deletion.includes('"' + name + '"') ||
      deletion.includes('"'+ name + '/"') ||
      deletion.includes('"' + name + '/'),
      "Missing account deletion collection reference: " + name);
    assert.ok(!found.has(name), "Duplicate inventory collection: " + name);
    found.add(name);
  }
  for (const [name] of outsideDeletion) {
    assert.match(firestoreRules, new RegExp("\\b" + name + "\\b"),
      "Business collection missing from Firestore rules: " + name);
    assert.ok(!found.has(name), "Collection has conflicting retention actions: " + name);
    found.add(name);
  }

  assert.match(privacy, /Profili\/account: fino alla disattivazione, poi fino a 12 mesi/i);
  assert.match(privacy, /Log tecnici:\s*90 giorni/i);
  assert.match(privacy, /Foto e allegati operativi:\s*24 mesi/i);
  assert.match(privacy, /documentazione amministrativo-contabile:\s*10 anni/i);
  assert.ok(retention.includes("const MAX_PER_COLLECTION = 10000;"),
    "Read-only retention audit cap unexpectedly changed");
  for (const name of ["activityLogs", "notificationEmailQueue", "gpsUpdateRequests",
    "appNotifications", "notifications"]) {
    assert.ok(retention.includes(name + ": {"),
      "Retention metadata classifier changed; re-review: " + name);
  }

  // Only these categories are currently within the protected retention
  // metadata scan; do not confuse them with all data held by the application.
  const protectedAuditCategories = [
    "activityLogs", "notificationEmailQueue", "gpsUpdateRequests",
    "appNotifications", "notifications"
  ];
  const report = {
    project: "VARGA CANTIERI",
    inventoryType: "source-code-only; no cloud data accessed",
    reviewStatus: "BLOCKED_PENDING_CONTROLLER_AND_RETENTION_REVIEW",
    demonstrablyCoveredDeletionCollections: deletionInventory.map(([collection, codeBehavior]) =>
      ({collection, codeBehavior})),
    explicitRetainedBusinessCollections: outsideDeletion.map(([collection, codeBehavior]) =>
      ({collection, codeBehavior})),
    protectedMetadataAuditCategories: protectedAuditCategories,
    policyTextDurations: {
      accountProfiles: "up to 12 months after deactivation (conditional)",
      technicalLogs: "90 days (conditional)",
      operationalAttachments: "24 months from work closure (conditional)",
      administrativeAccounting: "up to 10 years when justified"
    },
    unresolved,
    productionMutation: false,
    ordinaryDeletionEnabled: true
  };
  assert.equal(report.reviewStatus.startsWith("BLOCKED"), true);
  assert.ok(unresolved.length >= 7);
  assert.equal(new Set(report.demonstrablyCoveredDeletionCollections.map(x => x.collection)).size, 16);
  console.log(JSON.stringify(report, null, 2));
  console.log("PASS: source-level inventory consistent; privacy approval is NOT complete.");
}
audit();

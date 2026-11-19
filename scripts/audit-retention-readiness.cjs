"use strict";

/**
 * Review-only retention inventory. This script NEVER deletes, writes, or
 * exports document contents. It does not determine GDPR legal bases.
 * Default mode: synthetic tests, no Firebase. Production inventory requires
 * a dedicated protected workflow, exact project and explicit opt-in.
 *
 * These are conservative POSSIBLE review candidates, NOT deletion decisions.
 * Chat closure, attachments, Drive exports and backups require separate audits.
 */
const assert = require("node:assert/strict");
const PROJECT = "hera-app-6cd2b";
const MAX_PER_COLLECTION = 10000;
const CATEGORIES = Object.freeze({
  activityLogs: { anchor: "createdAt", days: 90 },
  notificationEmailQueue: { anchor: "createdAt", days: 30,
    terminal: ["sent", "delivered", "failed", "error", "cancelled", "cancelled_account_deleted"] },
  gpsUpdateRequests: { anchor: "resolvedAt", days: 30,
    terminal: ["approved", "rejected", "closed", "resolved"] },
  appNotifications: { anchor: "createdAt", days: 90, purpose: "transient" },
  notifications: { anchor: "createdAt", days: 90, purpose: "transient" }
});
const OUTCOMES = ["hold", "purpose-unclassified", "not-terminal", "missing-valid-anchor",
  "not-yet-due", "possible-manual-review"];

function dateFrom(value) {
  if (value && typeof value.toDate === "function") value = value.toDate();
  if (value instanceof Date) return Number.isFinite(value.getTime()) ? value : null;
  if (typeof value === "string" && /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?Z$/.test(value)) {
    const d = new Date(value);
    return Number.isFinite(d.getTime()) ? d : null;
  }
  return null;
}

function classify(category, raw, now = new Date()) {
  const policy = CATEGORIES[category];
  assert.ok(policy, "Unknown review category");
  const row = raw && typeof raw === "object" ? raw : {};
  const time = dateFrom(now);
  assert.ok(time, "Invalid audit clock");
  if (row.legalHold === true || row.retentionException === true) return "hold";
  // Work notifications must be classified before any time-based review.
  if (policy.purpose && row.retentionPurpose !== policy.purpose) return "purpose-unclassified";
  if (policy.terminal && !policy.terminal.includes(row.status)) return "not-terminal";
  const anchor = dateFrom(row[policy.anchor]);
  if (!anchor || anchor > time) return "missing-valid-anchor";
  return (time.getTime() - anchor.getTime()) < policy.days * 86400000
    ? "not-yet-due" : "possible-manual-review";
}

function emptyCounts() { return Object.fromEntries(OUTCOMES.map(key => [key, 0])); }
function summarise(category, rows, now) {
  const counts = emptyCounts();
  let total = 0;
  for (const row of rows) { counts[classify(category, row, now)]++; total++; }
  return { category, total, ...counts };
}

function selfTest() {
  const now = new Date("2026-11-19T12:00:00Z");
  const cases = [
    ["activityLogs", { createdAt: "2026-08-01T00:00:00Z" }, "possible-manual-review"],
    ["activityLogs", { createdAt: "2026-11-01T00:00:00Z" }, "not-yet-due"],
    ["activityLogs", { createdAt: "2025-01-01T00:00:00Z", legalHold: true }, "hold"],
    ["activityLogs", { createdAt: "not-a-date" }, "missing-valid-anchor"],
    ["notificationEmailQueue", { status: "pending", createdAt: "2025-01-01T00:00:00Z" }, "not-terminal"],
    ["notificationEmailQueue", { status: "sent", createdAt: "2026-09-01T00:00:00Z" }, "possible-manual-review"],
    ["gpsUpdateRequests", { status: "approved", createdAt: "2025-01-01T00:00:00Z" }, "missing-valid-anchor"],
    ["gpsUpdateRequests", { status: "approved", resolvedAt: "2026-10-01T00:00:00Z" }, "possible-manual-review"],
    ["notifications", { createdAt: "2025-01-01T00:00:00Z" }, "purpose-unclassified"],
    ["notifications", { createdAt: "2025-01-01T00:00:00Z", retentionPurpose: "transient" }, "possible-manual-review"],
    ["appNotifications", { createdAt: "2025-01-01T00:00:00Z", retentionPurpose: "transient", retentionException: true }, "hold"],
  ];
  for (const [kind, row, want] of cases) assert.equal(classify(kind, row, now), want);
  const batch = summarise("activityLogs", cases.filter(c => c[0] === "activityLogs").map(c => c[1]), now);
  assert.equal(batch.total, 4);
  assert.equal(batch["possible-manual-review"], 1);
  assert.equal(batch.hold, 1);
  assert.equal(classify("activityLogs", { createdAt: "2024-01-01T00:00:00Z" }, now), "possible-manual-review");
  console.log("PASS: retention review classifications, strict anchors, pending/holds, no Firebase access.");
}

async function liveAudit() {
  assert.equal(process.env.ALLOW_READ_ONLY_RETENTION_AUDIT, "1",
    "Live retention audit is disabled until approval.");
  const credentials = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  assert.ok(credentials, "Protected service account required");
  const fs = require("node:fs");
  const serviceAccount = JSON.parse(fs.readFileSync(credentials, "utf8"));
  assert.equal(serviceAccount.project_id, PROJECT, "Wrong Firebase project");
  const admin = require("../functions/node_modules/firebase-admin");
  admin.initializeApp({ credential: admin.credential.cert(serviceAccount), projectId: PROJECT });
  try {
    const db = admin.firestore();
    const report = [];
    for (const [category, policy] of Object.entries(CATEGORIES)) {
      const fields = [...new Set([policy.anchor, "status", "retentionPurpose", "legalHold", "retentionException"])];
      const col = db.collection(category);
      const counts = emptyCounts();
      let cursor = null, total = 0;
      while (true) {
        let query = col.orderBy(admin.firestore.FieldPath.documentId()).select(...fields).limit(250);
        if (cursor) query = query.startAfter(cursor);
        const snapshot = await query.get();
        for (const doc of snapshot.docs) {
          counts[classify(category, doc.data())]++;
          total++;
        }
        if (total > MAX_PER_COLLECTION) throw new Error("Row limit reached; stop and review scope");
        if (snapshot.size < 250) break;
        cursor = snapshot.docs[snapshot.docs.length - 1];
      }
      report.push({ category, total, ...counts });
    }
    // Aggregated counts ONLY: no document ID, coordinates, email or text.
    console.log(JSON.stringify({ project: PROJECT, retentionAudit: "review-only",
      automatedDeletion: false, categories: report }, null, 2));
  } finally {
    await admin.app().delete();
  }
}

if (require.main === module) {
  if (process.argv.includes("--self-test")) selfTest();
  else if (process.argv.includes("--live")) liveAudit().catch(err => {
    // Do not print raw SDK errors: they can contain document references.
    console.error("Retention read-only audit stopped:", err.code || "blocked-or-error");
    process.exitCode = 1;
  });
  else { console.error("Use --self-test or --live. No data accessed."); process.exitCode = 1; }
}
module.exports = { classify, summarise, CATEGORIES };

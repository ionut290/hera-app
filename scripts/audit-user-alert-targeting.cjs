"use strict";

/**
 * Read-only audit of legacy userAlerts delivery metadata.
 *
 * This is a release blocker, NOT an automatic migration. An alert with only
 * targetMemberNames cannot be safely converted into Firebase UIDs from a name:
 * duplicates and renamed workers could expose sensitive medical/absence data.
 *
 * Default: refuses access to production. --self-test uses synthetic data only.
 * Explicit live read-only inspection requires credentials, exact project
 * and ALLOW_READ_ONLY_USER_ALERT_AUDIT=1; prints counts, never message bodies,
 * worker names, email addresses or document IDs.
 */
const assert = require("node:assert/strict");

const PROJECT = "hera-app-6cd2b";
function presentIds(value) {
  return Array.isArray(value) ? value.filter(v => typeof v === "string" && !!v.trim()) : [];
}
function classification(value) {
  const data = value && typeof value === "object" ? value : {};
  const global = data.sendToAllRegistered === true;
  const members = presentIds(data.targetMemberNames);
  const ids = presentIds(data.targetUserIds);
  const single = typeof data.targetUserId === "string" && !!data.targetUserId.trim();
  const invalid = (data.targetUserIds != null && !Array.isArray(data.targetUserIds))
    || (data.targetMemberNames != null && !Array.isArray(data.targetMemberNames))
    || (data.sendToAllRegistered != null && typeof data.sendToAllRegistered !== "boolean");
  if (invalid) return "malformed";
  if (global) return "broadcast";
  if (ids.length || single) return members.length ? "explicit_with_legacy_names" : "explicit_uids";
  if (members.length) return "unmapped_legacy_names";
  return "unresolved";
}
const counters = [
  "broadcast", "explicit_uids", "explicit_with_legacy_names",
  "unmapped_legacy_names", "unresolved", "malformed"
];
function audit(docs) {
  const result = Object.fromEntries(counters.map(key => [key, 0]));
  let total = 0;
  for (const doc of docs) {
    result[classification(doc)]++;
    total++;
  }
  return { total, ...result,
    needsReview: result.unmapped_legacy_names + result.unresolved + result.malformed };
}
function selfTest() {
  const cases = [
    [{sendToAllRegistered: true, message: "secret"}, "broadcast"],
    [{targetUserIds: ["a"], message: "secret"}, "explicit_uids"],
    [{targetUserId: "a"}, "explicit_uids"],
    [{targetMemberNames: ["Worker A"]}, "unmapped_legacy_names"],
    [{targetUserIds: ["a"], targetMemberNames: ["Worker A"]}, "explicit_with_legacy_names"],
    [{targetMemberNames: [], targetUserIds: []}, "unresolved"],
    [{targetUserIds: "a"}, "malformed"],
    [{sendToAllRegistered: "true"}, "malformed"],
    [{source: "calendar-absence", targetMemberNames: ["Worker A"]}, "unmapped_legacy_names"],
  ];
  for (const [doc, expected] of cases) assert.equal(classification(doc), expected);
  const r = audit(cases.map(c => c[0]));
  assert.equal(r.total, cases.length);
  assert.equal(r.needsReview, 5);
  assert.equal(r.broadcast, 1);
  assert.equal(r.explicit_uids, 2);
  console.log("PASS: userAlerts recipient audit classifies legacy, malformed, broadcast and UID records with synthetic fixtures.");
}
async function liveReadOnly() {
  assert.equal(process.env.ALLOW_READ_ONLY_USER_ALERT_AUDIT, "1",
    "Production metadata audit is disabled by default.");
  const file = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  assert.ok(file, "A reviewed Google service-account credential file is required.");
  const fs = require("node:fs");
  const account = JSON.parse(fs.readFileSync(file, "utf8"));
  assert.equal(account.project_id, PROJECT, "Unexpected Firebase project: refusing audit.");
  // Only load the Admin SDK after the explicit read-only gate.
  const admin = require("../functions/node_modules/firebase-admin");
  admin.initializeApp({credential: admin.credential.cert(account), projectId: PROJECT});
  const col = admin.firestore().collection("userAlerts");
  const summary = Object.fromEntries(counters.map(key => [key, 0]));
  let total = 0;
  let last = null;
  do {
    let q = col.orderBy(admin.firestore.FieldPath.documentId()).limit(250);
    if (last) q = q.startAfter(last);
    const snapshot = await q.get();
    if (snapshot.empty) break;
    for (const doc of snapshot.docs) {
      // Select targeting metadata only; never print document values or IDs.
      const data = doc.data();
      summary[classification(data)]++;
      total++;
    }
    last = snapshot.docs.at(-1);
    if (snapshot.size < 250) break;
    assert.ok(total <= 100000, "Too many rows; manual review required.");
  } while (true);
  const result = { total, ...summary,
    needsReview: summary.unmapped_legacy_names + summary.unresolved + summary.malformed };
  console.log(JSON.stringify(result, null, 2));
  if (result.needsReview > 0) {
    console.error("BLOCK: unresolved targeting. Do not enforce private userAlerts rules yet.");
    process.exitCode = 2;
  }
}
if (process.argv.includes("--self-test")) selfTest();
else if (process.argv.includes("--live")) {
  liveReadOnly().catch(e => { console.error("Recipient audit aborted:", e.message); process.exitCode = 1; });
} else {
  console.error("Specify --self-test or --live. No Firebase data accessed.");
  process.exitCode = 1;
}

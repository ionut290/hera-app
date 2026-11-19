"use strict";

/**
 * Read-only PRELIMINARY gate before destructive App Review testing.
 *
 * This tool never creates, deletes or updates any Firebase resource. A clean
 * result is NOT permission to deploy account deletion or a full GDPR audit.
 *
 * A trusted Firebase administrator must explicitly mark the disposable user
 * with customClaims.appReviewDemo === true; self-assigned profile fields cannot
 * prove that an account is synthetic. Inspect only exact UID references, and
 * never print UIDs, emails, record identifiers, names, or document bodies.
 */
const assert = require("node:assert/strict");
const fs = require("node:fs");
const PROJECT = "hera-app-6cd2b";

const PROBES = Object.freeze([
  { collection: "personale", field: "linkedUserId", kind: "==" },
  { collection: "documents", field: "ownerUserId", kind: "==" },
  { collection: "chatMessages", field: "senderId", kind: "==" },
  { collection: "chatMessages", field: "recipientId", kind: "==" },
  { collection: "appNotifications", field: "createdByUid", kind: "==" },
  { collection: "notifications", field: "actorId", kind: "==" },
  { collection: "notifications", field: "recipientUserIds", kind: "array-contains" },
  { collection: "userAlerts", field: "targetUserId", kind: "==" },
  { collection: "userAlerts", field: "targetUserIds", kind: "array-contains" },
  { collection: "operatorCommessaHistory", field: "userId", kind: "==" },
  { collection: "oreReports", field: "userId", kind: "==" },
  { collection: "gpsUpdateRequests", field: "operatorId", kind: "==" },
  { collection: "notificationEmailQueue", field: "createdBy", kind: "==" },
  { collection: "activityLogs", field: "userId", kind: "==" }
]);
const PERSONAL_DOCS = Object.freeze(["privateDocuments", "operatorPositions", "userOperatorLinks"]);

function validUid(uid) {
  return typeof uid === "string" && /^[A-Za-z0-9_-]{1,128}$/.test(uid);
}
function hasTrustedDemoClaim(user) {
  const claims = user?.customClaims || {};
  // Both claims are signed by Firebase Auth Admin. apple_review denies
  // company Firestore/Storage access under the deployed isolation rules.
  // appReviewDemo distinguishes a new disposable test identity from an
  // existing reviewer account used by Apple to evaluate the normal app.
  return claims.appReviewDemo === true && claims.apple_review === true
    && claims.admin !== true && claims.isAdmin !== true
    && user?.disabled !== true;
}
function assess(user, links) {
  const linked = links.filter(item => item.found).map(item => item.category);
  return {
    syntheticClaimPresent: hasTrustedDemoClaim(user),
    checkedLinkCategories: links.length,
    associationsDetected: linked.length,
    categoriesRequiringReview: [...new Set(linked)].sort(),
    preliminaryGatePassed: hasTrustedDemoClaim(user) && linked.length === 0,
    grantsDeletionAuthorization: false,
    comprehensiveDataInventory: false
  };
}
function selfTest() {
  assert.ok(validUid("review-demo_42"));
  for (const uid of ["", "a.b", "../admin", "with@email", "has space", "a".repeat(129)]) {
    assert.equal(validUid(uid), false);
  }
  const clean = assess({ customClaims: { appReviewDemo: true, apple_review: true } }, [
    { category: "personale", found: false },
    { category: "documents", found: false }
  ]);
  assert.equal(clean.preliminaryGatePassed, true);
  assert.equal(clean.grantsDeletionAuthorization, false);
  assert.equal(clean.comprehensiveDataInventory, false);
  assert.equal(assess({ displayName: "appReviewDemo" }, []).preliminaryGatePassed, false);
  assert.equal(assess({ customClaims: { appReviewDemo: true } }, []).preliminaryGatePassed, false,
    "Demo-only marker must not bypass the reviewer isolation claim");
  assert.equal(assess({ customClaims: { apple_review: true } }, []).preliminaryGatePassed, false,
    "Existing Apple reviewer must never be reused for destructive demo");
  assert.equal(assess({ customClaims: { appReviewDemo: true, apple_review: true, admin: true } }, []).preliminaryGatePassed, false,
    "Administrator account cannot pass demo preflight");
  assert.equal(assess({ customClaims: { appReviewDemo: true, apple_review: true }, disabled: true }, []).preliminaryGatePassed, false);
  const linked = assess({ customClaims: { appReviewDemo: true, apple_review: true } }, [
    { category: "personale", found: true },
    { category: "personale", found: true },
    { category: "chatMessages", found: false }
  ]);
  assert.equal(linked.preliminaryGatePassed, false);
  assert.deepEqual(linked.categoriesRequiringReview, ["personale"]);
  console.log("PASS: isolated demo gate requires two signed isolation claims, rejects work links, never authorizes deletion.");
}

async function live() {
  assert.equal(process.env.ALLOW_READ_ONLY_DEMO_AUDIT, "1", "Live mode requires protected approval");
  assert.ok(validUid(process.env.REVIEW_DEMO_UID), "Valid isolated demo UID required");
  const credentials = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  assert.ok(credentials, "Dedicated protected Firebase credentials required");
  const service = JSON.parse(fs.readFileSync(credentials, "utf8"));
  assert.equal(service.project_id, PROJECT, "Wrong Firebase project");
  const admin = require("../functions/node_modules/firebase-admin");
  admin.initializeApp({ credential: admin.credential.cert(service), projectId: PROJECT });
  try {
    const uid = process.env.REVIEW_DEMO_UID;
    const user = await admin.auth().getUser(uid);
    // Fail closed before doing any collection reads if the user is not
    // explicitly marked by an administrator through trusted Auth custom claims.
    if (!hasTrustedDemoClaim(user)) {
      console.log(JSON.stringify({
        syntheticClaimPresent: false, preliminaryGatePassed: false,
        grantsDeletionAuthorization: false, comprehensiveDataInventory: false
      }));
      process.exitCode = 2;
      return;
    }
    const db = admin.firestore();
    const links = [];
    for (const probe of PROBES) {
      const snapshot = await db.collection(probe.collection)
        .where(probe.field, probe.kind, uid)
        .select(probe.field)
        .limit(1).get();
      links.push({ category: probe.collection, found: !snapshot.empty });
    }
    for (const collection of PERSONAL_DOCS) {
      const snapshot = await db.doc(collection + "/" + uid).get();
      // Existence only: never use or emit the doc's contents.
      links.push({ category: collection, found: snapshot.exists });
    }
    const report = assess(user, links);
    console.log(JSON.stringify(report, null, 2));
    if (!report.preliminaryGatePassed) process.exitCode = 2;
  } finally {
    await admin.app().delete();
  }
}
if (require.main === module) {
  if (process.argv.includes("--self-test")) selfTest();
  else if (process.argv.includes("--live")) live().catch(error => {
    // Avoid exposing user names, email or Firestore identifiers through errors.
    console.error("Demo readiness audit blocked:", error.code || "read-error");
    process.exitCode = 1;
  });
  else {
    console.error("Select --self-test or protected --live. No Firebase access attempted.");
    process.exitCode = 1;
  }
}
module.exports = { assess, validUid, hasTrustedDemoClaim, PROBES, PERSONAL_DOCS };

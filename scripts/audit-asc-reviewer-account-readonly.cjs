"use strict";
/**
 * Exact-account, READ-ONLY readiness check for the login listed in ASC.
 *
 * --self-test: no network, credentials or Firebase module required.
 * --inspect: deliberately gated, using a separately approved local read-only
 *            service credential. Reads exactly one Auth user and its own
 *            platformUsers/{uid} doc; never lists accounts or company data.
 *
 * No account creation, claim changes, writes, deletes, sign-in or deploy.
 */
const assert = require("node:assert/strict");
const fs = require("node:fs");
const EXPECTED_PROJECT = "hera-app-6cd2b";
const EXPECTED_EMAIL = "aaziz.aziz@operatori.vargacantieri.app";
const EMAIL = email => typeof email === "string" ? email.trim().toLowerCase() : "";

function inspectReviewer(user, doc) {
  const data = doc?.exists ? doc.data() : null;
  const claims = user?.customClaims || {};
  const claimKeys = Object.keys(claims);
  const checks = Object.freeze({
    accountExists: Boolean(user),
    enabled: Boolean(user && !user.disabled),
    verifiedEmail: user?.emailVerified === true,
    exactIdentity: EMAIL(user?.email) === EXPECTED_EMAIL,
    reviewerClaimOnly: claims.apple_review === true &&
      claimKeys.length === 1 && claimKeys[0] === "apple_review",
    ownProfileExists: Boolean(doc?.exists),
    profileIdentityMatches: Boolean(data && data.uid === user?.uid &&
      EMAIL(data.email) === EXPECTED_EMAIL),
    reviewerProfileOnly: Boolean(data && data.role === "apple_review" &&
      data.ruolo === "apple_review" && data.reviewer === true),
    nonAdmin: Boolean(data && data.admin === false &&
      data.isAdmin === false && data.banned === false),
    activeProfile: Boolean(data && data.statoAccount === "attivo" &&
      data.accountStatus === "attivo"),
    noOperationalAssignments: Boolean(data &&
      !Object.keys(data).some(k => [
        "teamId", "squadraId", "commessaIds", "permissions", "permittedCommesse",
        "allowedCommesse", "assignedCommesse", "operatorId", "personaleId"
      ].includes(k) && data[k] !== null && data[k] !== "" &&
        !(Array.isArray(data[k]) && data[k].length === 0) &&
        !(typeof data[k] === "object" && !Array.isArray(data[k]) &&
          data[k] && Object.keys(data[k]).length === 0))),
  });
  return { ready: Object.values(checks).every(Boolean), checks };
}
async function selfTest() {
  const user = {uid:"synthetic-reviewer",email:EXPECTED_EMAIL,disabled:false,
    emailVerified:true,customClaims:{apple_review:true}};
  const profile = {
    uid:user.uid,email:EXPECTED_EMAIL,role:"apple_review",ruolo:"apple_review",
    reviewer:true,admin:false,isAdmin:false,banned:false,
    statoAccount:"attivo",accountStatus:"attivo"
  };
  const snap = data => ({exists:true,data:()=>data});
  assert.equal(inspectReviewer(user,snap(profile)).ready,true);
  const cases = [
    [null,snap(profile)], [{...user,disabled:true},snap(profile)],
    [{...user,emailVerified:false},snap(profile)],
    [{...user,customClaims:{}},snap(profile)],
    [{...user,customClaims:{apple_review:true,admin:true}},snap(profile)],
    [user,{exists:false}],
    [user,snap({...profile,role:"user"})],
    [user,snap({...profile,ruolo:"user"})],
    [user,snap({...profile,reviewer:false})],
    [user,snap({...profile,admin:true})],
    [user,snap({...profile,banned:true})],
    [user,snap({...profile,accountStatus:"in_attesa"})],
    [user,snap({...profile,uid:"another-user"})],
    [user,snap({...profile,teamId:"real-team"})],
    [user,snap({...profile,permissions:{canViewPayroll:true}})]
  ];
  for (const [u,p] of cases) {
    assert.equal(inspectReviewer(u,p).ready,false);
  }
  // Read-only Firebase Console observation of 2026-11-20: active but no
  // reviewer role/flags, i.e. cannot be certified for safe Apple review.
  const observed = {...profile};
  delete observed.role;
  delete observed.ruolo;
  delete observed.reviewer;
  delete observed.isAdmin;
  assert.equal(inspectReviewer(user,snap(observed)).ready,false);
  console.log("PASS: 16 isolated synthetic reviewer conditions, including current observed missing role flags; no live network.");
}
async function inspect() {
  assert.equal(process.env.ALLOW_REVIEWER_READONLY_AUDIT, "1",
    "Explicit approval for live READ-only check is required");
  const credentialPath = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  assert.ok(credentialPath && fs.existsSync(credentialPath),
    "Review-approved credentials required");
  const creds = JSON.parse(fs.readFileSync(credentialPath, "utf8"));
  assert.equal(creds.project_id,EXPECTED_PROJECT,"Wrong Firebase project");
  assert.equal(EMAIL(process.env.REVIEWER_LOGIN_EMAIL),EXPECTED_EMAIL,
    "Only the exact ASC reviewer identity is permitted");
  const admin = require("../functions/node_modules/firebase-admin");
  const app = admin.initializeApp({
    credential:admin.credential.cert(creds),projectId:EXPECTED_PROJECT
  });
  try {
    let user;
    try {
      user = await admin.auth().getUserByEmail(EXPECTED_EMAIL);
    } catch (error) {
      if (error?.code !== "auth/user-not-found") throw error;
      console.log(JSON.stringify({ready:false,checks:{accountExists:false}}));
      process.exitCode=1;
      return;
    }
    const doc = await admin.firestore().collection("platformUsers")
      .doc(user.uid).get();
    const result = inspectReviewer(user, doc);
    // Print only named booleans. Never print IDs, documents, roles, passwords,
    // emails, service credentials, personal data or arbitrary server errors.
    console.log(JSON.stringify(result));
    if (!result.ready) process.exitCode=1;
  } finally {
    await app.delete();
  }
}
if (require.main === module) {
  const task = process.argv[2] === "--self-test" ? selfTest() :
    process.argv[2] === "--inspect" ? inspect() :
      Promise.reject(Error("Select --self-test or --inspect"));
  Promise.resolve(task).catch(error => {
    console.error("Reviewer readiness check blocked:",error.code || "guard-failed");
    process.exitCode=1;
  });
}
module.exports = { inspectReviewer };

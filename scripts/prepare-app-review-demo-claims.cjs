"use strict";

// Operates on ONE approved, disabled and newly created Firebase Auth demo.
// --inspect is read-only; --apply changes only two Auth custom claims.
// Never enables, creates or deletes accounts or deploys any Firebase resource.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const crypto = require("node:crypto");
const PROJECT = "hera-app-6cd2b";
const APPROVED_EMAIL_HASH = "458f589f8bba0bbe8dcc14d2f4e45efc5a5c3b135c2df4530413ebfdf52fa1da";
const CREATED_AFTER = Date.parse("2026-11-18T23:00:00Z");
const EXPECTED_CLAIMS = Object.freeze({ apple_review: true, appReviewDemo: true });
const FIELDS = [
  ["personale","linkedUserId","uid"], ["personale","email","email"],
  ["documents","ownerUserId","uid"],
  ["chatMessages","senderId","uid"], ["chatMessages","recipientId","uid"],
  ["chatMessages","senderEmail","email"],
  ["appNotifications","createdByUid","uid"],
  ["notifications","actorId","uid"], ["notifications","recipientUserIds","uid","array-contains"],
  ["userAlerts","userId","uid"], ["oreReports","userId","uid"],
  ["operatorCommessaHistory","userId","uid"], ["activityLogs","userId","uid"],
  ["activityLogs","userEmail","email"], ["notificationEmailQueue","createdBy","uid"],
  ["notificationEmailQueue","to","email"]
];
const UID_DOCS = ["platformUsers","privateDocuments","operatorPositions","userOperatorLinks"];
const hash = email => crypto.createHash("sha256").update(email).digest("hex");

function validateTarget(user, rawEmail) {
  const email = String(rawEmail || "").trim().toLowerCase();
  assert.equal(hash(email), APPROVED_EMAIL_HASH, "Unapproved email alias");
  assert.equal(user?.email?.toLowerCase(), email, "Wrong Firebase identity");
  assert.equal(user?.disabled, true, "The test account must remain DISABLED");
  assert.match(user?.uid || "", /^[A-Za-z0-9_-]{1,128}$/);
  assert.ok(Date.parse(user?.metadata?.creationTime || "") >= CREATED_AFTER,
    "Not the newly created disposable demo");
  assert.ok((user?.providerData || []).every(x => x.providerId === "password"),
    "Federated identity detected");
  const previous = user.customClaims || {};
  const names = Object.keys(previous);
  assert.ok(names.length === 0 ||
    (names.length === 2 && previous.apple_review === true && previous.appReviewDemo === true),
    "Unexpected prior role/custom claims");
  return { email, uid: user.uid, alreadySet: names.length === 2 };
}

async function assertNoWorkLinks(db, uid, email) {
  let checks = 0;
  for (const [col,field,identity,op = "=="] of FIELDS) {
    const value = identity === "uid" ? uid : email;
    const matches = await db.collection(col).where(field,op,value).select(field).limit(1).get();
    assert.equal(matches.empty,true,"Associated data found: claims NOT changed");
    checks++;
  }
  for (const col of UID_DOCS) {
    const doc = await db.doc(col + "/" + uid).get();
    assert.equal(doc.exists,false,"Existing account profile requires review");
    checks++;
  }
  return checks;
}

async function selfTest() {
  const email = "sqadrahera+varga-demo@gmail.com";
  const base = { email, uid: "new-demo", disabled:true, providerData:[],
    metadata:{ creationTime:"2026-11-19T12:00:00Z" }};
  assert.equal(validateTarget(base,email).alreadySet,false);
  assert.equal(validateTarget({...base,customClaims:EXPECTED_CLAIMS},email).alreadySet,true);
  for (const bad of [
    {...base,disabled:false}, {...base,email:"other@example.com"},
    {...base,metadata:{creationTime:"2026-11-01T00:00:00Z"}},
    {...base,customClaims:{admin:true}},
    {...base,customClaims:{apple_review:true}},
    {...base,providerData:[{providerId:"google.com"}]}
  ]) assert.throws(() => validateTarget(bad,email));
  assert.throws(() => validateTarget(base,"squadrahera+varga-demo@gmail.com"));
  const db = {
    collection:()=>({where:()=>({select:()=>({limit:()=>({get:async()=>({empty:true})})})})}),
    doc:()=>({get:async()=>({exists:false})})
  };
  assert.equal(await assertNoWorkLinks(db,base.uid,email),FIELDS.length+UID_DOCS.length);
  console.log("PASS: protected demo claims only for exact disabled, new, unlinked non-admin identity.");
}

async function live(mode) {
  assert.equal(process.env.ALLOW_PROTECTED_DEMO_CLAIM_PREPARATION,"1");
  if(mode==="apply") assert.equal(process.env.ALLOW_DEMO_CLAIM_WRITE,"1");
  const credentialPath = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  assert.ok(credentialPath,"Protected Firebase credential required");
  const cred = JSON.parse(fs.readFileSync(credentialPath,"utf8"));
  assert.equal(cred.project_id,PROJECT,"Wrong Firebase project");
  const email = String(process.env.REVIEW_DEMO_EMAIL || "").trim().toLowerCase();
  assert.equal(hash(email),APPROVED_EMAIL_HASH,"Not the approved demo alias");
  const admin = require("../functions/node_modules/firebase-admin");
  const app = admin.initializeApp({credential:admin.credential.cert(cred),projectId:PROJECT});
  try {
    const before = await admin.auth().getUserByEmail(email);
    const target = validateTarget(before,email);
    const checks = await assertNoWorkLinks(admin.firestore(),target.uid,target.email);
    if(mode==="apply" && !target.alreadySet) {
      await admin.auth().setCustomUserClaims(target.uid,EXPECTED_CLAIMS);
    }
    const after = await admin.auth().getUser(target.uid);
    assert.equal(after.disabled,true,"Disabled state changed");
    assert.equal(after.email?.toLowerCase(),target.email,"Email changed");
    if(mode==="apply") assert.deepEqual(after.customClaims,EXPECTED_CLAIMS);
    console.log(JSON.stringify({mode,disabled:true,workLinkChecks:checks,
      claimsReady:after.customClaims?.apple_review===true&&after.customClaims?.appReviewDemo===true,
      accountChanged:mode==="apply"&&!target.alreadySet, accountDeleted:false}));
  } finally { await app.delete(); }
}
if(require.main===module) {
  const arg = process.argv[2];
  const task = arg==="--self-test" ? selfTest() :
    arg==="--inspect" ? live("inspect") :
    arg==="--apply" ? live("apply") : Promise.reject(Error("Invalid mode"));
  Promise.resolve(task).catch(error=>{
    // Never print personal data, UIDs, credentials or document contents.
    console.error("Demo claims gate blocked:",error.code||"guard-failed");
    process.exitCode=1;
  });
}
module.exports={validateTarget,assertNoWorkLinks};

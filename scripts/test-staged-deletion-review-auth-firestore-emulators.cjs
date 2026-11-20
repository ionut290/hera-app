"use strict";
/*
 * REAL Auth + Firestore emulators, actual staged callable handler bodies.
 * NO network, remote Firebase credentials, production data or deletions.
 */
const assert=require("node:assert/strict");
const fs=require("node:fs");
const vm=require("node:vm");
const project="demo-deletion-review-e2e";
for(const key of ["FIRESTORE_EMULATOR_HOST","FIREBASE_AUTH_EMULATOR_HOST"]){
  const v=process.env[key]||"";
  if(!/^(127\.0\.0\.1|localhost|\[::1\]):[0-9]{2,5}$/.test(v)){
    throw Error("Both Auth and Firestore emulators must be on loopback: "+key);
  }
}
for(const key of ["GCLOUD_PROJECT","GOOGLE_CLOUD_PROJECT","FIREBASE_PROJECT"]){
  if(process.env[key]&&process.env[key]!==project)throw Error("Wrong demo-only project "+key);
}
if(process.env.GOOGLE_APPLICATION_CREDENTIALS)throw Error("Service accounts forbidden in this test");
const entry=fs.readFileSync("functions/main.js","utf8");
const normal=fs.readFileSync(".github/workflows/deploy-firebase-functions.yml","utf8");
assert.ok(!entry.includes("account-deletion-review-staged"));
assert.ok(!entry.includes('require("./account-deletion-request")'));
assert.ok(!normal.includes("functions:reviewVargaAccountDeletionRequest"));

const admin=require("../functions/node_modules/firebase-admin");
admin.initializeApp({projectId:project});
const db=admin.firestore();
const auth=admin.auth();
class HttpsError extends Error{constructor(code,message){super(message);this.code=code;}}
function load(path){
  const targets={};
  // Load actual source in the host JS realm (Firestore validates Promise
  // instanceof against its own realm in transaction callbacks).
  const factory=vm.runInThisContext("(function (exports, require, console) {\n"+
    fs.readFileSync(path,"utf8")+"\n})",{filename:path});
  factory(targets,(name)=>{
    if(name==="firebase-admin")return admin;
    if(name==="firebase-functions/v2/https")return{
      HttpsError,
      onCall(opts,handler){
        assert.equal(opts.region,"europe-west1");
        assert.ok(opts.maxInstances<=3);
        return handler;
      }
    };
    throw Error("Unexpected production dependency "+name);
  },{error(){}});
  return targets;
}
const intake=load("functions/account-deletion-request.js");
const review=load("functions/account-deletion-review-staged.js");
assert.deepEqual(Object.keys(intake).sort(),
  ["getOwnVargaAccountDeletionStatus","requestOwnVargaAccountDeletion"].sort());
assert.deepEqual(Object.keys(review).sort(),
  ["listVargaAccountDeletionReviewQueue","reviewVargaAccountDeletionRequest"].sort());
const pid="synthetic-"+process.pid;
const reviewerUid=pid+"-reviewer",workerUid=pid+"-alice",otherUid=pid+"-bob";
const users=[reviewerUid,workerUid,otherUid];
const now=Math.floor(Date.now()/1000);
const workerRequest=uid=>({auth:{uid,token:{auth_time:now}},data:{confirmation:"RICHIEDO ELIMINAZIONE"}});
const token=(extra={})=>({
  uid:reviewerUid,
  token:{admin:true,privacy_deletion_reviewer:true,auth_time:now,...extra}
});
const call=(data,authContext=token())=>({auth:authContext,data});
const rejects=(promise,code)=>assert.rejects(promise,e=>e instanceof HttpsError&&e.code===code);
async function cleanup(){
  for(const uid of users){
    await Promise.all([
      db.doc("accountDeletionRequests/"+uid).delete(),
      db.doc("platformUsers/"+uid).delete(),
      db.doc("privateDocuments/"+uid).delete()
    ]);
    try{await auth.deleteUser(uid);}catch(e){if(e.code!=="auth/user-not-found")throw e;}
  }
  await admin.app().delete();
}
(async()=>{
  try{
    // Create actual local Auth user records, no email or personal information.
    await Promise.all(users.map(uid=>auth.createUser({uid})));
    await auth.setCustomUserClaims(reviewerUid,{admin:true,privacy_deletion_reviewer:true});
    const business={synthetic:true,retainedBusiness:"KEEP"};
    await db.doc("platformUsers/"+workerUid).set(business);
    await db.doc("privateDocuments/"+workerUid).set({synthetic:true,ownerUserId:workerUid});

    // Genuine intake requests persist as Firestore server timestamps, not a stub.
    const first=await intake.requestOwnVargaAccountDeletion(workerRequest(workerUid));
    const second=await intake.requestOwnVargaAccountDeletion(workerRequest(otherUid));
    assert.equal(first.status,"pending");assert.equal(second.status,"pending");
    const queue=await review.listVargaAccountDeletionReviewQueue(call({}));
    assert.equal(queue.requests.length,2);
    assert.deepEqual(queue.requests.map(x=>x.uid).sort(),[workerUid,otherUid].sort());
    for(const item of queue.requests)assert.deepEqual(Object.keys(item).sort(),
      ["uid","status","reopened"].sort(),"Response must exclude email and sensitive data");

    // A worker or forged review token cannot see or change other identities.
    await rejects(review.listVargaAccountDeletionReviewQueue(
      call({}, {uid:workerUid,token:{auth_time:now}})),"permission-denied");
    await rejects(review.listVargaAccountDeletionReviewQueue(
      call({}, {...token(),uid:workerUid})),"permission-denied");
    await rejects(review.listVargaAccountDeletionReviewQueue(
      call({},token({auth_time:now-601}))),"failed-precondition");
    await rejects(review.reviewVargaAccountDeletionRequest(
      call({uid:workerUid,action:"complete"})),"invalid-argument");
    assert.equal((await db.doc("accountDeletionRequests/"+workerUid).get()).data().status,"pending");

    // True Firestore transaction handles two competing reviewers atomically:
    // only one may claim the same pending request.
    const attempts=await Promise.allSettled([0,1].map(()=>
      review.reviewVargaAccountDeletionRequest(call({uid:workerUid,action:"start_review"}))
    ));
    assert.equal(attempts.filter(x=>x.status==="fulfilled").length,1);
    assert.equal(attempts.filter(x=>x.status==="rejected"&&x.reason?.code==="failed-precondition").length,1);
    assert.equal((await db.doc("accountDeletionRequests/"+workerUid).get()).data().status,"in_review");

    // Revoking live claims denies a previously authorized unexpired token.
    await auth.setCustomUserClaims(reviewerUid,{admin:true});
    await rejects(review.reviewVargaAccountDeletionRequest(
      call({uid:workerUid,action:"reject",reasonCode:"identity_not_verified"})),"permission-denied");
    await auth.setCustomUserClaims(reviewerUid,{admin:true,privacy_deletion_reviewer:true});
    await auth.updateUser(reviewerUid,{disabled:true});
    await rejects(review.listVargaAccountDeletionReviewQueue(call({})),"permission-denied");
    await auth.updateUser(reviewerUid,{disabled:false});

    await rejects(review.reviewVargaAccountDeletionRequest(call({
      uid:workerUid,action:"reject",reasonCode:"free_text_with_personal_info"
    })),"invalid-argument");
    const rejected=await review.reviewVargaAccountDeletionRequest(call({
      uid:workerUid,action:"reject",reasonCode:"retention_review_required"
    }));
    assert.equal(rejected.status,"rejected");
    assert.equal((await db.doc("accountDeletionRequests/"+workerUid).get()).data().rejectionReasonCode,
      "retention_review_required");
    // The rejected user is allowed to reopen the request without erasure.
    const reopened=await intake.requestOwnVargaAccountDeletion(workerRequest(workerUid));
    assert.equal(reopened.status,"pending");assert.equal(reopened.reopened,true);
    const updated=(await db.doc("accountDeletionRequests/"+workerUid).get()).data();
    assert.equal(updated.reopenCount,1);
    assert.equal(updated.rejectionReasonCode,"retention_review_required");
    assert.equal(typeof updated.reopenedAt.toDate,"function");
    assert.equal((await intake.getOwnVargaAccountDeletionStatus(
      {auth:{uid:workerUid,token:{auth_time:now}},data:{}})).status,"pending");
    // Completed records cannot be set by these operators; other data untouched.
    await rejects(review.reviewVargaAccountDeletionRequest(
      call({uid:otherUid,action:"reject",reasonCode:"identity_not_verified"})),"failed-precondition");
    assert.equal((await db.doc("accountDeletionRequests/"+otherUid).get()).data().status,"pending");
    assert.deepEqual((await db.doc("platformUsers/"+workerUid).get()).data(),business);
    assert.equal((await db.doc("privateDocuments/"+workerUid).get()).data().ownerUserId,workerUid);
    for(const uid of users) assert.equal((await auth.getUser(uid)).uid,uid);

    console.log("PASS: local Auth+Firestore emulators — claims/revocation, queue isolation, transactional concurrency, reject/reopen, no deleted users or business records");
  }finally{await cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1;});

"use strict";
// Calls ACTUAL staged function handlers against REAL LOCAL Firestore transactions.
// All Firebase auth claims are synthetic; no remote services or credentials.
const assert=require("node:assert/strict");
const fs=require("node:fs");
const vm=require("node:vm");
const project="demo-account-deletion-request";
const host=process.env.FIRESTORE_EMULATOR_HOST || "";
if(!/^(?:127\.0\.0\.1|localhost|\[::1\]):[0-9]{2,5}$/.test(host)) throw Error("Only loopback Firestore emulator permitted");
for(const key of ["GCLOUD_PROJECT","GOOGLE_CLOUD_PROJECT"])
  if(process.env[key] && process.env[key]!==project) throw Error("Wrong demo project");
if(process.env.GOOGLE_APPLICATION_CREDENTIALS) throw Error("No service account credentials allowed");
const admin=require("../functions/node_modules/firebase-admin");
admin.initializeApp({projectId:project});
const db=admin.firestore();
class HttpsError extends Error { constructor(code,message){super(message);this.code=code;} }
const exportsForTest={};
// Compile in the same JS realm as the real Firestore Admin SDK: the SDK
// checks that a transaction callback returns an instanceof Promise.
// vm.runInNewContext would create a foreign-realm Promise and reject it.
const loadHandlers=vm.runInThisContext(
  "(function(exports, require, console) {\\n" +
  fs.readFileSync("functions/account-deletion-request.js","utf8") + "\\n})",
  {filename:"functions/account-deletion-request.js"}
);
loadHandlers(exportsForTest, function(name){
  if(name==="firebase-admin")return admin;
  if(name==="firebase-functions/v2/https")return {
    HttpsError,onCall(opts,fn){assert.equal(opts.region,"europe-west1");return fn;}
  };
  throw Error("Unexpected dependency");
}, console);
assert.deepEqual(Object.keys(exportsForTest).sort(),["requestOwnVargaAccountDeletion","getOwnVargaAccountDeletionStatus"].sort());
const ask=exportsForTest.requestOwnVargaAccountDeletion,read=exportsForTest.getOwnVargaAccountDeletionStatus;
const firstUid="synthetic-"+process.pid+"-alice", secondUid="synthetic-"+process.pid+"-bob";
const now=Math.floor(Date.now()/1000);
const auth=(uid,age=0)=>({uid,token:{auth_time:now-age}});
const request=(uid,age=0)=>({auth:auth(uid,age),data:{confirmation:"RICHIEDO ELIMINAZIONE"}});
const rejects=(p,code)=>assert.rejects(p,e=>e instanceof HttpsError&&e.code===code);
(async()=>{
  const profile={synthetic:true,tag:"preserve"}, privateDoc={ownerUserId:firstUid,synthetic:true};
  await db.doc("platformUsers/"+firstUid).set(profile);
  await db.doc("privateDocuments/"+firstUid).set(privateDoc);
  await rejects(ask({auth:null,data:request(firstUid).data}),"unauthenticated");
  await rejects(read({auth:null,data:{}}),"unauthenticated");
  await rejects(ask(request(firstUid,601)),"failed-precondition");
  await rejects(ask({auth:auth(firstUid),data:{confirmation:"RICHIEDO ELIMINAZIONE",uid:secondUid}}),"invalid-argument");
  await rejects(read({auth:auth(firstUid),data:{uid:secondUid}}),"invalid-argument");
  assert.equal((await read({auth:auth(firstUid),data:{}})).requested,false);
  const first=await ask(request(firstUid));
  assert.equal(first.status,"pending");
  assert.equal(first.alreadyExists,false);
  assert.equal((await ask(request(firstUid))).alreadyExists,true);
  assert.equal((await read({auth:auth(firstUid),data:{}})).status,"pending");
  assert.equal((await read({auth:auth(secondUid),data:{}})).requested,false);
  const saved=(await db.doc("accountDeletionRequests/"+firstUid).get()).data();
  assert.deepEqual(Object.keys(saved).sort(),["uid","status","submittedAt","origin","reviewBeforeErasureRequired"].sort());
  assert.equal(saved.uid,firstUid);
  assert.equal(saved.origin,"authenticated_app");
  assert.equal(saved.reviewBeforeErasureRequired,true);
  assert.equal(typeof saved.submittedAt.toDate,"function");
  // Operator update is synthetic and isolated to emulator, NOT a real fulfillment flow.
  await db.doc("accountDeletionRequests/"+firstUid).update({status:"in_review"});
  assert.equal((await ask(request(firstUid))).status,"in_review");
  assert.equal((await read({auth:auth(firstUid),data:{}})).status,"in_review");
  assert.equal((await ask(request(secondUid))).status,"pending");
  assert.equal((await read({auth:auth(secondUid),data:{}})).status,"pending");
  assert.equal((await db.collection("accountDeletionRequests").get()).size,2);
  assert.deepEqual((await db.doc("platformUsers/"+firstUid).get()).data(),profile);
  assert.deepEqual((await db.doc("privateDocuments/"+firstUid).get()).data(),privateDoc);
  for(const collection of ["accountDeletionRequests","platformUsers","privateDocuments"])
    for(const uid of [firstUid,secondUid]) await db.doc(collection+"/"+uid).delete();
  await admin.app().delete();
  console.log("PASS: real local Firestore transactions, UID isolation, server timestamps, duplicate requests and preserved synthetic data");
})().catch(e=>{console.error(e);process.exitCode=1;admin.app().delete().catch(()=>{});});

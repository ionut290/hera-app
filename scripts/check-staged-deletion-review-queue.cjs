"use strict";
// Execute the REAL staged review handlers with synthetic Admin/Auth/Firestore.
// No Firebase credentials, real account, network calls, deletion or deployment.
const assert=require("node:assert/strict");
const fs=require("node:fs");
const vm=require("node:vm");
const path=require("node:path");
const file="functions/account-deletion-review-staged.js";
const source=fs.readFileSync(file,"utf8");
const main=fs.readFileSync("functions/main.js","utf8");
const routine=fs.readFileSync(".github/workflows/deploy-firebase-functions.yml","utf8");
assert.ok(!main.includes("account-deletion-review-staged"));
assert.ok(!routine.includes("functions:reviewVargaAccountDeletionRequest"));
assert.ok(!routine.includes("functions:listVargaAccountDeletionReviewQueue"));
assert.ok(!source.includes(".deleteUser("));
assert.ok(!source.includes("tx.delete("));
assert.ok(!source.includes('status: "completed"'));
const records=new Map([
  ["worker-1",{uid:"worker-1",status:"pending",reopenCount:1,sensitiveBusiness:"PRESERVE"}],
  ["worker-2",{uid:"worker-2",status:"in_review",sensitiveBusiness:"PRESERVE"}],
  ["worker-3",{uid:"worker-3",status:"completed",sensitiveBusiness:"PRESERVE"}],
  ["worker-4",{uid:"worker-4",status:"rejected",sensitiveBusiness:"PRESERVE"}]
]);
const writes=[];
const currentLive={disabled:false,customClaims:{admin:true,privacy_deletion_reviewer:true}};
class HttpsError extends Error {constructor(code,message){super(message);this.code=code;}}
const admin={
  auth(){return {async getUser(uid){if(uid!=="reviewer-1")throw Error("Unknown admin");return currentLive;}};},
  firestore:Object.assign(()=>({
    collection(name){
      assert.equal(name,"accountDeletionRequests");
      return {
        doc(uid){return {uid,async get(){return {exists:records.has(uid),data:()=>records.get(uid)};}};},
        where(field,op,values){
          assert.equal(field,"status");assert.equal(op,"in");
          assert.deepEqual(Array.from(values),["pending","in_review"]);
          return {limit(num){assert.equal(num,25);
            return {async get(){
              const docs=[...records.entries()].filter(([uid,d])=>values.includes(d.status)).slice(0,num)
                .map(([id,d])=>({id,data:()=>d}));
              return {docs,size:docs.length};
            }};
          }};
        }
      };
    },
    async runTransaction(work){
      return work({
        async get(ref){return ref.get();},
        update(ref,patch){
          if(!records.has(ref.uid))throw Error("Missing ref");
          records.set(ref.uid,{...records.get(ref.uid),...patch});
          writes.push({uid:ref.uid,patch});
        }
      });
    }
  }),{FieldValue:{serverTimestamp:()=>"<synthetic-server-time>"}})
};
const handlers={};
vm.runInNewContext(source,{
  require(name){
    if(name==="firebase-admin")return admin;
    if(name==="firebase-functions/v2/https")return {
      HttpsError,
      onCall(options,fn){
        assert.equal(options.region,"europe-west1");
        assert.equal(options.invoker,"public");
        assert.ok(options.maxInstances<=2);
        return fn;
      }
    };
    throw Error("Unexpected module "+name);
  },exports:handlers,Date,Math,Number,Set,Object,Array,String,
  console:{error(){}}
},{filename:file});
const queue=handlers.listVargaAccountDeletionReviewQueue;
const review=handlers.reviewVargaAccountDeletionRequest;
assert.equal(typeof queue,"function");
assert.equal(typeof review,"function");
assert.deepEqual(Object.keys(handlers).sort(),
  ["listVargaAccountDeletionReviewQueue","reviewVargaAccountDeletionRequest"].sort());
const now=Math.floor(Date.now()/1000);
const token=(extras={})=>({
  uid:"reviewer-1",
  token:{admin:true,privacy_deletion_reviewer:true,auth_time:now,...extras}
});
const caller=(data, extras={})=>({auth:token(extras),data});
const rejects=(promise,code)=>assert.rejects(promise,e=>e instanceof HttpsError&&e.code===code);
(async()=>{
  await rejects(queue({auth:null,data:{}}),"permission-denied");
  await rejects(queue(caller({uid:"worker-1"})),"invalid-argument");
  await rejects(queue(caller({}, {privacy_deletion_reviewer:false})),"permission-denied");
  await rejects(queue(caller({}, {admin:false})),"permission-denied");
  await rejects(queue(caller({}, {auth_time:now-601})),"failed-precondition");
  currentLive.customClaims.privacy_deletion_reviewer=false;
  await rejects(queue(caller({})),"permission-denied");
  currentLive.customClaims.privacy_deletion_reviewer=true;
  currentLive.disabled=true;
  await rejects(queue(caller({})),"permission-denied");
  currentLive.disabled=false;

  const data=await queue(caller({}));
  assert.equal(data.requests.length,2);
  assert.deepEqual(Array.from(data.requests,x=>x.uid),["worker-1","worker-2"]);
  assert.equal(data.requests[0].reopened,true);
  assert.equal(data.requests[1].reopened,false);
  assert.deepEqual(Object.keys(data.requests[0]).sort(),["reopened","status","uid"].sort());
  assert.equal(writes.length,0);

  await rejects(review(caller({uid:"worker-1",action:"complete"})),"invalid-argument");
  await rejects(review(caller({uid:"worker-1",action:"start_review",force:true})),"invalid-argument");
  await rejects(review(caller({uid:"reviewer-1",action:"start_review"})),"invalid-argument");
  await rejects(review(caller({uid:"worker/1",action:"start_review"})),"invalid-argument");
  await rejects(review(caller({uid:"worker-3",action:"start_review"})),"failed-precondition");
  await rejects(review(caller({uid:"worker-4",action:"reject",reasonCode:"identity_not_verified"})),"failed-precondition");
  await rejects(review(caller({uid:"worker-1",action:"reject",reasonCode:"identity_not_verified"})),"failed-precondition");
  await rejects(review(caller({uid:"worker-2",action:"reject",reasonCode:"free text with PII"})),"invalid-argument");
  await rejects(review(caller({uid:"worker-2",action:"reject"})),"invalid-argument");
  assert.equal(writes.length,0);

  const first=await review(caller({uid:"worker-1",action:"start_review"}));
  assert.equal(first.status,"in_review");
  assert.equal(writes.length,1);
  assert.equal(records.get("worker-1").reviewedByUid,"reviewer-1");
  assert.equal(records.get("worker-1").reviewedAt,"<synthetic-server-time>");
  assert.equal(records.get("worker-1").sensitiveBusiness,"PRESERVE");
  await rejects(review(caller({uid:"worker-1",action:"start_review"})),"failed-precondition");
  const second=await review(caller({uid:"worker-1",action:"reject",reasonCode:"retention_review_required"}));
  assert.equal(second.status,"rejected");
  assert.equal(writes.length,2);
  assert.equal(records.get("worker-1").rejectionReasonCode,"retention_review_required");
  assert.equal(records.get("worker-1").sensitiveBusiness,"PRESERVE");
  assert.deepEqual(Array.from(data.requests,x=>x.status),["pending","in_review"],
    "Earlier queue snapshot should remain immutable");
  assert.equal(records.get("worker-2").status,"in_review");
  assert.equal(records.get("worker-3").status,"completed");
  assert.equal(records.get("worker-4").status,"rejected");
  assert.equal((await queue(caller({}))).requests.length,1);
  console.log("PASS: live+token dual-claim check, fresh session, PII-minimal queue, strict review/reject transitions, no completion, no deletion, preserved other accounts.");
})().catch(e=>{console.error(e);process.exitCode=1;});

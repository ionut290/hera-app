"use strict";

// Exact user-only non-destructive Firebase Admin audit. Never writes, sends
// credentials, prints UID or exposes personal fields. Runs in approved CI only.
const assert=require("node:assert/strict");
const fs=require("node:fs");
const crypto=require("node:crypto");
const PROJECT="hera-app-6cd2b";
const EMAIL="aaziz.aziz@operatori.vargacantieri.app";
const HASH=crypto.createHash("sha256").update(EMAIL).digest("hex");
const hash=s=>crypto.createHash("sha256").update(String(s||"").trim().toLowerCase()).digest("hex");
function summary(record, profile, adminList) {
  assert.equal(hash(record?.email),HASH,"Wrong Firebase identity");
  const data=profile?.exists?profile.data():null;
  const status=data?.statoAccount??data?.accountStatus??null;
  const role=data?.role??data?.ruolo??null;
  const bans=data?.banned===true;
  const isAdminList=Array.isArray(adminList)&&adminList.some(x=>hash(x)===HASH);
  const claims=record?.customClaims||{};
  const adminLike=Boolean(claims.admin===true||claims.isAdmin===true||
    isAdminList||String(role||"").toLowerCase()==="admin");
  return {
    exists:true,enabled:record.disabled===false,
    emailVerified:record.emailVerified===true,
    providerTypes:(record.providerData||[]).map(x=>x.providerId).filter(x=>["password","google.com","apple.com"].includes(x)),
    platformProfileExists:Boolean(data),
    profileStatus:typeof status==="string"?status:"unknown",
    profileRole:typeof role==="string"?role:"unknown",
    banned:bans,adminLike,
    reviewClaim:Boolean(claims.apple_review===true),
    // Sensitive affiliations and UID are deliberately not output.
    readyForApple:false,
    note:"This report does not authorize third-party access or verify credentials."
  };
}
async function main() {
  assert.equal(process.env.ALLOW_SCOPED_REVIEW_ACCOUNT_AUDIT,"1");
  const cred=JSON.parse(fs.readFileSync(process.env.GOOGLE_APPLICATION_CREDENTIALS,"utf8"));
  assert.equal(cred.project_id,PROJECT);
  const admin=require("../functions/node_modules/firebase-admin");
  const app=admin.initializeApp({credential:admin.credential.cert(cred),projectId:PROJECT});
  try {
    const record=await admin.auth().getUserByEmail(EMAIL);
    assert.equal(hash(record.email),HASH);
    const db=admin.firestore();
    const [profile,adminDoc]=await Promise.all([
      db.doc("platformUsers/"+record.uid).get(),
      db.doc("appConfig/adminUsers").get()
    ]);
    const list=adminDoc.exists?adminDoc.data()?.emails:[];
    console.log("READ_ONLY_AUDIT="+JSON.stringify(summary(record,profile,list)));
  } finally { await app.delete(); }
}
async function test() {
  const record={email:EMAIL,disabled:false,emailVerified:true,
    providerData:[{providerId:"password"}],customClaims:{}};
  const profile={exists:true,data:()=>({statoAccount:"attivo",role:"user",banned:false})};
  const out=summary(record,profile,[]);
  assert.equal(out.profileStatus,"attivo");assert.equal(out.adminLike,false);
  assert.equal(out.readyForApple,false);assert.equal(out.platformProfileExists,true);
  assert.throws(()=>summary({...record,email:"other@example.test"},profile,[]));
  assert.equal(summary({...record,customClaims:{admin:true}},profile,[]).adminLike,true);
  assert.ok(!Object.keys(out).includes("uid"));
  console.log("PASS: chosen-account audit stays read-only and never declares App Review ready.");
}
if(require.main===module) Promise.resolve(process.argv[2]==="--self-test"?test():
  process.argv[2]==="--audit"?main():Promise.reject(Error("Unknown mode")))
  .catch(()=>{console.error("Read-only audit unavailable; no account data changed.");process.exitCode=1;});

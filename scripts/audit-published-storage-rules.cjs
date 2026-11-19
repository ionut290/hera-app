"use strict";
// Inspect the deployed Firebase Security Rules RELEASES, not the console editor.
// READ-ONLY Google API: GET projects/*/releases and GET projects/*/rulesets.
// Never download customer files, inspect Auth users, or print rules contents.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const crypto = require("node:crypto");
const PROJECT = "hera-app-6cd2b";
const STORAGE_BUCKET = "hera-app-6cd2b.firebasestorage.app";
const API = "https://firebaserules.googleapis.com/v1/";
const sha256 = s => crypto.createHash("sha256").update(s).digest("hex");

function assessStorageSource(code, expected) {
  const sections = [
    { name:"documents", marker:"match /documents/{ownerUserId}/{documentId}/{fileName}" },
    { name:"legacy", marker:"match /lavori-occasionali/{plantId}/{ownerUserId}/{documentId}/{fileName}" }
  ];
  const gates = {};
  for (let i=0;i<sections.length;i++) {
    const here = code.indexOf(sections[i].marker);
    const next = i+1<sections.length ? code.indexOf(sections[i+1].marker) : code.length;
    const part = here>=0 && (next>here) ? code.slice(here,next) : "";
    const allows = [...part.matchAll(/\ballow\s+(read|write)\s*:\s*if\s+([\s\S]*?);/g)];
    gates[sections[i].name] = here>=0 && allows.length===2 &&
      new Set(allows.map(m=>m[1])).size===2 &&
      allows.every(m=>m[2].includes("!isAppleReviewAccount()") && m[2].includes("request.auth != null"));
  }
  const matches = [...code.matchAll(/\bmatch\s+\//g)].length;
  const allows = [...code.matchAll(/\ballow\s+/g)].length;
  return {
    matchesRepository: sha256(code) === sha256(expected),
    signedAppleReviewClaim: /request\.auth\.token\.get\("apple_review",\s*false\)\s*==\s*true/.test(code),
    documentsDeniedToReviewers: gates.documents,
    legacyDeniedToReviewers: gates.legacy,
    expectedPathCountOnly: matches===3 && allows===4,
    reviewRequired: !(gates.documents && gates.legacy && matches===3 && allows===4),
    sourceHashPrefix: sha256(code).slice(0,16)
  };
}
function selfTest() {
  const good = fs.readFileSync("storage.rules","utf8");
  const check=assessStorageSource(good,good);
  assert.equal(check.matchesRepository,true);
  assert.equal(check.signedAppleReviewClaim,true);
  assert.equal(check.documentsDeniedToReviewers,true);
  assert.equal(check.legacyDeniedToReviewers,true);
  assert.equal(check.expectedPathCountOnly,true);
  assert.equal(check.reviewRequired,false);
  const broken=good.replace("!isAppleReviewAccount() && (\n", "(");
  assert.equal(assessStorageSource(broken,good).documentsDeniedToReviewers,false);
  const extra=good.replace("  match /b/{bucket}/o {","  match /b/{bucket}/o {\n    match /{allPaths=**} { allow read, write: if request.auth != null; }\n");
  assert.equal(assessStorageSource(extra,good).reviewRequired,true);
  console.log("PASS: ruleset auditor fails closed on missing reviewer guard and unexpected catch-all.");
}
async function live() {
  assert.equal(process.env.APPROVE_PUBLISHED_RULES_READ_ONLY,"1","Reviewer-only live read");
  const credentials=process.env.GOOGLE_APPLICATION_CREDENTIALS;
  assert.ok(credentials,"Protected credential required");
  const account=JSON.parse(fs.readFileSync(credentials,"utf8"));
  assert.equal(account.project_id,PROJECT,"Wrong Firebase project");
  const {GoogleAuth}=require("../functions/node_modules/google-auth-library");
  const auth=new GoogleAuth({keyFile:credentials,scopes:["https://www.googleapis.com/auth/cloud-platform"]});
  const client=await auth.getClient();
  const apiGet=async path=>{
    assert.ok(path.startsWith("projects/"+PROJECT+"/"),"Wrong project path");
    const response=await client.request({url:API+path,method:"GET"});
    return response.data;
  };
  const releases=[];
  let pageToken="";
  for(let i=0;i<20;i++){
    const url="projects/"+PROJECT+"/releases?pageSize=100" +
      (pageToken ? "&pageToken="+encodeURIComponent(pageToken) : "");
    const response=await apiGet(url);
    for(const release of response.releases || []) releases.push(release);
    pageToken=response.nextPageToken || "";
    if(!pageToken) break;
    if(i===19) throw Error("Release inventory pagination exceeded");
  }
  const storage=releases.filter(r=>r.name?.startsWith("projects/"+PROJECT+"/releases/firebase.storage"));
  const firestore=releases.filter(r=>r.name?.startsWith("projects/"+PROJECT+"/releases/cloud.firestore"));
  const expected=fs.readFileSync("storage.rules","utf8");
  const storageResults=[];
  for(const rel of storage) {
    assert.match(rel.rulesetName || "",/^projects\/hera-app-6cd2b\/rulesets\/[A-Za-z0-9_-]+$/);
    const ruleset=await apiGet(rel.rulesetName);
    const sources=(ruleset?.source?.files || []).map(x=>x.content).filter(x=>typeof x==="string");
    assert.ok(sources.length,"Published Storage ruleset has no readable source");
    const code=sources.join("\n");
    storageResults.push({
      release:rel.name.slice(("projects/"+PROJECT+"/releases/").length),
      updatedAt:rel.updateTime || null,
      ...assessStorageSource(code,expected)
    });
  }
  const result={
    project:PROJECT,
    readOnly:true,
    releaseCount:releases.length,
    storageReleaseCount:storage.length,
    firestoreReleaseCount:firestore.length,
    storageRules:storageResults,
    storageActiveVerified:storageResults.length>0,
    authorizationToDeploy:false,
    productionDataRead:false
  };
  console.log(JSON.stringify(result,null,2));
  if(!result.storageActiveVerified ||
    storageResults.some(x=>x.reviewRequired || !x.signedAppleReviewClaim)) process.exitCode=2;
}
if(require.main===module){
  const arg=process.argv[2];
  const work=arg==="--self-test"?selfTest():
    arg==="--live"?live():Promise.reject(Error("Select --self-test or --live"));
  Promise.resolve(work).catch(err=>{
    console.error("Published Rules API audit blocked:",err.code||"read-failed");
    process.exitCode=1;
  });
}
module.exports={assessStorageSource};

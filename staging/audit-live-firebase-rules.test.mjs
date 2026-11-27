import test from 'node:test';
import assert from 'node:assert/strict';
import {getActiveRules} from './audit-live-firebase-rules.mjs';
import {assembleRules} from './assemble-candidate-rules.mjs';
const release='projects/varga-cantieri-v2-staging/releases/cloud.firestore';
const ruleset='projects/varga-cantieri-v2-staging/rulesets/test-id';
function mock({releaseName=release,rulesetName=ruleset,source=assembleRules(),status=200}={}){
 let calls=0;
 return async(url,opt)=>{
  assert.equal(opt.method,'GET');assert.equal(opt.redirect,'error');
  assert.equal(opt.headers.Authorization,'Bearer synthetic-only');
  assert.equal(new URL(url).origin,'https://firebaserules.googleapis.com');
  assert.ok(++calls<=2);
  if(calls===1){
   assert.equal(url,'https://firebaserules.googleapis.com/v1/'+release);
   return {ok:status===200,status,json:async()=>({name:releaseName,rulesetName,updateTime:'2026-11-26T12:00:00Z'})};
  }
  assert.equal(url,'https://firebaserules.googleapis.com/v1/'+ruleset);
  return {ok:true,json:async()=>({name:ruleset,source:{files:[{name:'firestore.rules',content:source}]}})};
 };
}
test('exact candidate comparison without exposing rules',async()=>{
 const r=await getActiveRules({token:'synthetic-only',request:mock()});
 assert.equal(r.matchesTestedCandidate,true);assert.match(r.sha256,/^[a-f0-9]{64}$/);assert.equal(r.rulesetName,ruleset);assert.equal('content' in r,false);
});
test('detect changed rules',async()=>assert.equal((await getActiveRules({token:'synthetic-only',request:mock({source:'// different\nrules_version = "2";'})})).matchesTestedCandidate,false));
test('reject production ruleset before second request',async()=>assert.rejects(getActiveRules({token:'synthetic-only',request:mock({rulesetName:'projects/hera-app-6cd2b/rulesets/test-id'})}),/identity/));
test('reject unexpected release',async()=>assert.rejects(getActiveRules({token:'synthetic-only',request:mock({releaseName:'other'})}),/identity/));
test('fail closed on denied API',async()=>assert.rejects(getActiveRules({token:'synthetic-only',request:mock({status:403})}),/HTTP 403/));
test('reject empty source',async()=>assert.rejects(getActiveRules({token:'synthetic-only',request:mock({source:''})}),/Empty/));

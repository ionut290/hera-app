'use strict';
const assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto');
const {createLocalAdapter}=require('./auth-adapter.emulator.cjs');
const {assertStagingEnvironment}=require('../../functions/multi-tenant/staging-adapter.cjs');
const valid={VARGA_MULTI_TENANT_STAGING:'1',GOOGLE_CLOUD_PROJECT:'varga-cantieri-v2-staging'};
assert.doesNotThrow(()=>assertStagingEnvironment(valid));
for(const env of [{},{...valid,GOOGLE_CLOUD_PROJECT:'hera-app-6cd2b'},{...valid,GCLOUD_PROJECT:'hera-app-6cd2b'},{...valid,FIREBASE_CONFIG:'{"projectId":"hera-app-6cd2b"}'},{...valid,FIRESTORE_EMULATOR_HOST:'127.0.0.1:8180'},{...valid,FIREBASE_AUTH_EMULATOR_HOST:'127.0.0.1:9199'}]){
  assert.throws(()=>assertStagingEnvironment(env),{code:'staging-only'});
}
const adapter=createLocalAdapter();
const {startLocalCompanyHttp}=require('./company-http.emulator.cjs');
let http;
const password='Demo-'+randomUUID();
const rejects=(fn,code)=>assert.rejects(fn,{code});
async function account(label,claims={},verified=true){
  const email=label+'-'+randomUUID()+'@example.invalid';
  const u=await adapter.auth.createUser({email,password,emailVerified:verified});
  await adapter.auth.setCustomUserClaims(u.uid,claims);
  const response=await fetch('http://127.0.0.1:9199/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=demo-key',{
    method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({email,password,returnSecureToken:true})
  });
  const data=await response.json();
  if(!response.ok||!data.idToken)throw Error('Synthetic sign-in failed');
  return {uid:u.uid,email,token:data.idToken};
}
async function main(){
  const admin=await account('platform',{platform_admin:true});
  const owner=await account('owner'), worker=await account('worker');
  const review=await account('review',{platform_admin:true,apple_review:true});
  const unverified=await account('unverified',{},false);
  const call=(u,action,payload)=>adapter.dispatch(u.token,action,payload);
  http=await startLocalCompanyHttp(adapter.dispatch);
  const viaHttp=async(token,action,payload={},origin=http.origin)=>{
    const r=await fetch(http.origin+'/api/v2/companies',{method:'POST',headers:{'content-type':'application/json','authorization':'Bearer '+token,origin},body:JSON.stringify({action,payload})});
    return {status:r.status,...await r.json()};
  };
  assert.equal((await viaHttp('invalid','listCompanies')).status,401);
  assert.equal((await viaHttp(worker.token,'listCompanies')).status,403);
  assert.equal((await viaHttp(admin.token,'listCompanies',{},'https://example.invalid')).status,403);
  assert.equal((await viaHttp(admin.token,'invite',{})).status,400);
  const httpOwner=await account('http-owner');
  const created=await viaHttp(admin.token,'createCompany',{name:'HTTP synthetic company',ownerEmail:httpOwner.email});
  assert.equal(created.status,200);
  assert.ok(created.result.tenantId);
  assert.equal((await adapter.db.collection('tenants/'+created.result.tenantId+'/commesse').get()).empty,true);
  const list=await viaHttp(admin.token,'listCompanies',{limit:100});
  assert.equal(list.status,200);
  assert.ok(list.result.companies.some(c=>c.id===created.result.tenantId));
  assert.equal((await viaHttp(admin.token,'setCompanyStatus',{tenantId:created.result.tenantId,status:'suspended',reason:'Synthetic HTTP test'})).status,200);
  assert.equal((await adapter.db.doc('tenants/'+created.result.tenantId).get()).data().status,'suspended');
  assert.equal((await viaHttp(admin.token,'setCompanyStatus',{tenantId:created.result.tenantId,status:'active',reason:'Synthetic HTTP test'})).status,200);
  await rejects(()=>adapter.dispatch('invalid-token','listCompanies'),'unauthenticated');
  await rejects(()=>call(review,'listCompanies'),'unauthenticated');
  await rejects(()=>call(unverified,'listCompanies'),'unauthenticated');
  await rejects(()=>call(worker,'createCompany',{name:'Forbidden',ownerEmail:owner.email,platform_admin:true,uid:admin.uid}),'forbidden');
  await rejects(()=>call(admin,'createCompany',{name:'Invalid owner',ownerEmail:unverified.email}),'invalid-owner');
  const {tenantId}=await call(admin,'createCompany',{name:'Synthetic verified company',ownerEmail:owner.email,ownerUid:worker.uid});
  const root='tenants/'+tenantId;
  assert.equal((await adapter.db.doc(root+'/members/'+owner.uid).get()).data().role,'owner');
  assert.equal((await adapter.db.doc(root+'/members/'+worker.uid).get()).exists,false);
  // Provision a second company through the real local Auth adapter, not mock UI.
  await adapter.db.doc(root+'/commesse/synthetic-existing').set({name:'Synthetic existing job'});
  const emptyOwner=await account('empty-owner');
  const emptyCompany=await call(admin,'createCompany',{name:'Synthetic empty company',ownerEmail:emptyOwner.email});
  assert.notEqual(emptyCompany.tenantId,tenantId);
  const emptyRoot='tenants/'+emptyCompany.tenantId;
  for(const module of ['commesse','personale','squadre','mezzi','neve_commesse','servizioNevePercorsi','documents','rapportini']){
    assert.equal((await adapter.db.collection(emptyRoot+'/'+module).limit(1).get()).empty,true,'New tenant unexpectedly contains '+module);
  }
  assert.equal((await adapter.db.doc(emptyRoot+'/members/'+emptyOwner.uid).get()).data().role,'owner');
  assert.equal((await adapter.db.doc(emptyRoot+'/members/'+owner.uid).get()).exists,false);
  assert.equal((await adapter.db.doc(root+'/commesse/synthetic-existing').get()).data().name,'Synthetic existing job');
  await rejects(()=>call(owner,'invite',{tenantId:emptyCompany.tenantId,recipientEmail:worker.email,role:'worker'}),'forbidden');
  const invitation=await call(owner,'invite',{tenantId,recipientEmail:worker.email,role:'worker'});
  await call(worker,'accept',{tenantId,inviteToken:invitation.token,role:'owner'});
  assert.equal((await adapter.db.doc(root+'/members/'+worker.uid).get()).data().role,'worker');
  await rejects(()=>call(worker,'listCompanies',{platform_admin:true}),'forbidden');
  assert.ok((await call(admin,'listCompanies',{limit:100})).companies.some(c=>c.id===tenantId));
  await call(admin,'setCompanyStatus',{tenantId,status:'suspended',reason:'Synthetic test'});
  await rejects(()=>call(owner,'invite',{tenantId,recipientEmail:worker.email,role:'worker'}),'forbidden');
  await call(admin,'setCompanyStatus',{tenantId,status:'active',reason:'Synthetic test'});
  // Workspace package: genuine company selection, role and tenant-scoped jobs.
  const second=await account('workspace-second');
  const secondCompany=await call(admin,'createCompany',{name:'Workspace B',ownerEmail:second.email});
  const firstCompany=await call(admin,'createCompany',{name:'Workspace A',ownerEmail:owner.email});
  const ar='tenants/'+firstCompany.tenantId,br='tenants/'+secondCompany.tenantId;
  await adapter.db.doc(ar+'/commesse/job-a').set({nome:'Ordinary A'});
  await adapter.db.doc(ar+'/commesse/job-a2').set({nome:'Ordinary A2'});
  await adapter.db.doc(ar+'/neve_commesse/snow-a').set({nome:'Snow A'});
  await adapter.db.doc(br+'/commesse/job-b').set({nome:'Ordinary B'});
  await adapter.db.doc(br+'/neve_commesse/snow-b').set({nome:'Snow B'});
  await adapter.db.doc('commesse/legacy-private').set({nome:'Legacy must never appear'});
  const my=await viaHttp(owner.token,'listMyCompanies');
  assert.equal(my.status,200);assert.ok(my.result.companies.some(c=>c.id===firstCompany.tenantId));assert.ok(!my.result.companies.some(c=>c.id===secondCompany.tenantId));
  assert.equal((await viaHttp(owner.token,'openCompany',{tenantId:secondCompany.tenantId})).status,403);
  assert.equal((await viaHttp(admin.token,'openCompany',{tenantId:firstCompany.tenantId})).status,403);
  const opened=await viaHttp(owner.token,'openCompany',{tenantId:firstCompany.tenantId});
  assert.equal(opened.result.company.role,'owner');assert.equal(opened.result.capabilities.readCommesse,true);
  const jobs=await viaHttp(owner.token,'listCompanyCommesse',{tenantId:firstCompany.tenantId,kind:'ordinary',limit:1});
  assert.equal(jobs.status,200);assert.deepEqual(jobs.result.commesse.map(x=>x.name),['Ordinary A']);assert.equal(jobs.result.nextCursor,'job-a');
  const page2=await viaHttp(owner.token,'listCompanyCommesse',{tenantId:firstCompany.tenantId,kind:'ordinary',limit:1,after:jobs.result.nextCursor});
  assert.deepEqual(page2.result.commesse.map(x=>x.name),['Ordinary A2']);assert.equal(page2.result.nextCursor,null);
  const snow=await viaHttp(owner.token,'listCompanyCommesse',{tenantId:firstCompany.tenantId,kind:'snow'});
  assert.deepEqual(snow.result.commesse.map(x=>x.name),['Snow A']);
  assert.equal((await viaHttp(second.token,'listCompanyCommesse',{tenantId:firstCompany.tenantId,kind:'snow'})).status,403);
  assert.equal((await viaHttp(owner.token,'listCompanyCommesse',{tenantId:secondCompany.tenantId,kind:'ordinary'})).status,403);
  assert.equal((await viaHttp(owner.token,'listCompanyCommesse',{tenantId:firstCompany.tenantId,kind:'../../commesse'})).status,400);
  const invitationB=await call(second,'invite',{tenantId:secondCompany.tenantId,recipientEmail:owner.email,role:'worker'});
  await call(owner,'accept',{tenantId:secondCompany.tenantId,inviteToken:invitationB.token});
  const both=await viaHttp(owner.token,'listMyCompanies');
  assert.ok(both.result.companies.some(c=>c.id===firstCompany.tenantId));assert.ok(both.result.companies.some(c=>c.id===secondCompany.tenantId));
  assert.equal((await viaHttp(owner.token,'openCompany',{tenantId:secondCompany.tenantId})).result.capabilities.readCommesse,false);
  assert.equal((await viaHttp(owner.token,'listCompanyCommesse',{tenantId:secondCompany.tenantId,kind:'snow',role:'owner'})).status,403);
  await call(admin,'setCompanyStatus',{tenantId:firstCompany.tenantId,status:'suspended',reason:'Synthetic workspace suspension'});
  assert.equal((await viaHttp(owner.token,'listCompanyCommesse',{tenantId:firstCompany.tenantId,kind:'ordinary'})).status,403);
  assert.ok(!(await viaHttp(owner.token,'listMyCompanies')).result.companies.some(c=>c.id===firstCompany.tenantId));
  await call(admin,'setCompanyStatus',{tenantId:firstCompany.tenantId,status:'active',reason:'Synthetic workspace reactivation'});
  await adapter.db.doc(br+'/members/'+owner.uid).update({status:'revoked'});
  assert.equal((await viaHttp(owner.token,'openCompany',{tenantId:secondCompany.tenantId})).status,403);
  assert.ok(!(await viaHttp(owner.token,'listMyCompanies')).result.companies.some(c=>c.id===secondCompany.tenantId));
  console.log('Company workspace and ordinary/snow isolation HTTP assertions passed');
  await adapter.auth.setCustomUserClaims(admin.uid,{});
  assert.equal((await viaHttp(admin.token,'listCompanies')).status,403);
  await rejects(()=>call(admin,'listCompanies'),'forbidden');
  await adapter.auth.updateUser(owner.uid,{disabled:true});
  await rejects(()=>call(owner,'invite',{tenantId,recipientEmail:worker.email,role:'worker'}),'unauthenticated');
  console.log('Firebase Auth emulator adapter integration passed');
}
main().catch(e=>{console.error(e.code||e.name);process.exitCode=1;}).finally(async()=>{if(http)await http.close();await adapter.close();});

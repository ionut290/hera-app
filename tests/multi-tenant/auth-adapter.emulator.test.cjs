'use strict';
const assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto');
const {createLocalAdapter}=require('./auth-adapter.emulator.cjs');
const adapter=createLocalAdapter();
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
  await rejects(()=>adapter.dispatch('invalid-token','listCompanies'),'unauthenticated');
  await rejects(()=>call(review,'listCompanies'),'unauthenticated');
  await rejects(()=>call(unverified,'listCompanies'),'unauthenticated');
  await rejects(()=>call(worker,'createCompany',{name:'Forbidden',ownerEmail:owner.email,platform_admin:true,uid:admin.uid}),'forbidden');
  await rejects(()=>call(admin,'createCompany',{name:'Invalid owner',ownerEmail:unverified.email}),'invalid-owner');
  const {tenantId}=await call(admin,'createCompany',{name:'Synthetic verified company',ownerEmail:owner.email,ownerUid:worker.uid});
  const root='tenants/'+tenantId;
  assert.equal((await adapter.db.doc(root+'/members/'+owner.uid).get()).data().role,'owner');
  assert.equal((await adapter.db.doc(root+'/members/'+worker.uid).get()).exists,false);
  const invitation=await call(owner,'invite',{tenantId,recipientEmail:worker.email,role:'worker'});
  await call(worker,'accept',{tenantId,inviteToken:invitation.token,role:'owner'});
  assert.equal((await adapter.db.doc(root+'/members/'+worker.uid).get()).data().role,'worker');
  await rejects(()=>call(worker,'listCompanies',{platform_admin:true}),'forbidden');
  assert.ok((await call(admin,'listCompanies',{limit:100})).companies.some(c=>c.id===tenantId));
  await call(admin,'setCompanyStatus',{tenantId,status:'suspended',reason:'Synthetic test'});
  await rejects(()=>call(owner,'invite',{tenantId,recipientEmail:worker.email,role:'worker'}),'forbidden');
  await call(admin,'setCompanyStatus',{tenantId,status:'active',reason:'Synthetic test'});
  await adapter.auth.setCustomUserClaims(admin.uid,{});
  await rejects(()=>call(admin,'listCompanies'),'forbidden');
  await adapter.auth.updateUser(owner.uid,{disabled:true});
  await rejects(()=>call(owner,'invite',{tenantId,recipientEmail:worker.email,role:'worker'}),'unauthenticated');
  console.log('Firebase Auth emulator adapter integration passed');
}
main().catch(e=>{console.error(e.code||e.name);process.exitCode=1;}).finally(()=>adapter.close());

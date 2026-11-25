'use strict';
// Synthetic Firestore transaction integration tests; emulator only.
const assert = require('node:assert/strict');
const { createHash } = require('node:crypto');
const { initializeApp, deleteApp } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');
const { makeTenantService } = require('./tenant-service.prototype.cjs');
const projectId = 'demo-varga-multi-tenant';
if (process.env.FIRESTORE_EMULATOR_HOST !== '127.0.0.1:8180') throw new Error('Local emulator required');
const app = initializeApp({projectId});
const db = getFirestore(app);
const store = {runTransaction: fn => db.runTransaction(async transaction => fn({
  get: async path => {const doc = await transaction.get(db.doc(path)); return doc.exists ? doc.data() : null;},
  create: (path,data) => transaction.create(db.doc(path),data),
  set: (path,data) => transaction.set(db.doc(path),data)
}))};
const auth = (uid,extra={})=>({uid,email:uid+'@example.invalid',email_verified:true,...extra});
const platform = auth('platform',{platform_admin:true}), alice=auth('owner-a'), bob=auth('owner-b'), worker=auth('worker'), second=auth('second'), admin=auth('admin');
let now=Date.now();
const service=makeTenantService(store,()=>now);
const rejects=(fn,code)=>assert.rejects(fn,{code});
async function main(){
  await rejects(()=>service.createCompany(worker,{name:'Forbidden',ownerUid:worker.uid,ownerEmail:worker.email}),'forbidden');
  await rejects(()=>service.createCompany(auth('platform',{platform_admin:true,apple_review:true}),{name:'Review',ownerUid:alice.uid,ownerEmail:alice.email}),'unauthenticated');
  const a=(await service.createCompany(platform,{name:'Synthetic A',ownerUid:alice.uid,ownerEmail:alice.email})).tenantId;
  const b=(await service.createCompany(platform,{name:'Synthetic B',ownerUid:bob.uid,ownerEmail:bob.email})).tenantId;
  assert.notEqual(a,b);
  assert.equal((await db.doc('tenants/'+a+'/members/'+alice.uid).get()).data().role,'owner');
  await rejects(()=>service.invite(alice,{tenantId:b,recipientEmail:worker.email,role:'worker'}),'forbidden');
  await rejects(()=>service.invite(platform,{tenantId:a,recipientEmail:worker.email,role:'worker'}),'forbidden');
  await rejects(()=>service.invite(alice,{tenantId:a,recipientEmail:worker.email,role:'owner'}),'invalid-role');
  const invite=await service.invite(alice,{tenantId:a,recipientEmail:worker.email,role:'worker'});
  const hash=createHash('sha256').update(invite.token).digest('hex');
  const stored=await db.doc('tenants/'+a+'/invites/'+hash).get();
  assert.equal(stored.data().status,'pending');
  assert.ok(!JSON.stringify(stored.data()).includes(invite.token));
  await rejects(()=>service.accept(second,{tenantId:a,token:invite.token}),'invalid-invite');
  await rejects(()=>service.accept(worker,{tenantId:b,token:invite.token}),'invalid-invite');
  await rejects(()=>service.accept(auth('worker',{email_verified:false}),{tenantId:a,token:invite.token}),'unauthenticated');
  const race=await Promise.allSettled([service.accept(worker,{tenantId:a,token:invite.token}),service.accept(worker,{tenantId:a,token:invite.token})]);
  assert.equal(race.filter(x=>x.status==='fulfilled').length,1);
  assert.equal(race.filter(x=>x.status==='rejected').length,1);
  assert.equal((await db.doc('tenants/'+a+'/members/'+worker.uid).get()).data().role,'worker');
  assert.equal((await db.doc('users/'+worker.uid+'/memberships/'+a).get()).data().status,'active');
  await rejects(()=>service.invite(worker,{tenantId:a,recipientEmail:second.email,role:'worker'}),'forbidden');
  await rejects(()=>service.accept(worker,{tenantId:a,token:invite.token}),'invalid-invite');
  const adminInvite=await service.invite(alice,{tenantId:a,recipientEmail:admin.email,role:'admin'});
  await service.accept(admin,{tenantId:a,token:adminInvite.token});
  await rejects(()=>service.invite(admin,{tenantId:a,recipientEmail:second.email,role:'admin'}),'forbidden');
  const adminIssued=await service.invite(admin,{tenantId:a,recipientEmail:second.email,role:'worker'});
  await db.doc('tenants/'+a+'/members/'+admin.uid).update({status:'revoked'});
  await rejects(()=>service.accept(second,{tenantId:a,token:adminIssued.token}),'invalid-invite');
  const expired=await service.invite(alice,{tenantId:a,recipientEmail:second.email,role:'worker'});
  now=expired.expiresAt;
  await rejects(()=>service.accept(second,{tenantId:a,token:expired.token}),'invalid-invite');
  const bInvite=await service.invite(bob,{tenantId:b,recipientEmail:worker.email,role:'client'});
  await service.accept(worker,{tenantId:b,token:bInvite.token});
  assert.equal((await db.doc('tenants/'+b+'/members/'+worker.uid).get()).data().role,'client');
  assert.equal((await db.doc('tenants/'+a+'/members/'+worker.uid).get()).data().role,'worker');
  const duplicate=await service.invite(alice,{tenantId:a,recipientEmail:worker.email,role:'manager'});
  await rejects(()=>service.accept(worker,{tenantId:a,token:duplicate.token}),'membership-exists');
  const suspended=await service.invite(bob,{tenantId:b,recipientEmail:second.email,role:'worker'});
  await db.doc('tenants/'+b).update({status:'suspended'});
  await rejects(()=>service.accept(second,{tenantId:b,token:suspended.token}),'invalid-invite');
  console.log('Tenant backend transaction integration tests passed');
}
main().catch(e=>{console.error(e);process.exitCode=1;}).finally(async()=>{await db.terminate();await deleteApp(app);});

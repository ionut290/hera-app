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
  // Complete registry/economic preview package, identical IDs across tenants.
  const moduleFixtures=[
    ['personnel','personale','neve_personale',{fullName:'Person A',email:'private@example.invalid'},'Person A'],
    ['vehicles','mezzi','neve_mezzi',{nId:'Vehicle A',targa:'SECRET'},'Vehicle A'],
    ['teams','squadreCommesse','neve_squadre',{commessaNome:'Team A',squadre:[{operatore:'private'}]},'Team A']
  ];
  for(const [module,ordinary,snowCollection,data,expected] of moduleFixtures){
    for(const [kind,collection] of [['ordinary',ordinary],['snow',snowCollection]]){
      await adapter.db.doc(ar+'/'+collection+'/same-id').set(data);
      await adapter.db.doc(br+'/'+collection+'/same-id').set({fullName:'Private B',nome:'Private B',commessaNome:'Private B'});
      await adapter.db.doc(collection+'/same-id').set({fullName:'Legacy private',nome:'Legacy private'});
      const payload={tenantId:firstCompany.tenantId,module,kind};
      const response=await viaHttp(owner.token,'listCompanyModule',payload);
      assert.equal(response.status,200);assert.deepEqual(response.result.rows.map(x=>x.name),[expected]);
      assert.ok(!JSON.stringify(response.result).includes('private@example.invalid'));
      assert.ok(!JSON.stringify(response.result).includes('SECRET'));
      assert.equal((await viaHttp(second.token,'listCompanyModule',payload)).status,403);
      assert.equal((await viaHttp(admin.token,'listCompanyModule',payload)).status,403);
      assert.equal((await viaHttp(owner.token,'listCompanyModule',{...payload,tenantId:secondCompany.tenantId,role:'owner'})).status,403);
      const empty=await viaHttp(emptyOwner.token,'listCompanyModule',{...payload,tenantId:emptyCompany.tenantId});
      assert.equal(empty.status,200);assert.deepEqual(empty.result.rows,[]);
    }
  }
  await adapter.db.doc(ar+'/sharedStaticViews/registri__corrente').set({payload:{personale:[{id:'p1',fullName:'Aggregated A',email:'hidden@example.invalid'}],mezzi:[{id:'m1',nId:'Aggregated vehicle A'}]}});
  const aggregate=await viaHttp(owner.token,'listCompanyModule',{tenantId:firstCompany.tenantId,module:'personnel',kind:'ordinary'});
  assert.deepEqual(aggregate.result.rows,[{id:'p1',name:'Aggregated A'}]);
  for(const [kind,collection,job] of [['ordinary','commesse','job-a'],['snow','neve_commesse','snow-a']]){
    const target=ar+'/'+collection+'/'+job;
    await adapter.db.doc(target+'/prezziario/item').set({codiceVoce:'P01',descrizione:'Price A',unitaMisura:'h',prezzoBase:30,prezzoRibassato:25});
    await adapter.db.doc(target+'/lavorazioni/item').set({denominazione:'Work A',codiceVocePrezzo:'P01',totale:50,stato:'FATTO',operatoreNome:'private'});
    // Same parent and item IDs in another company: no fallback.
    await adapter.db.doc(br+'/'+collection+'/'+job).set({nome:'B duplicate ID'});
    await adapter.db.doc(br+'/'+collection+'/'+job+'/prezziario/item').set({descrizione:'Price B',prezzoBase:999});
    for(const [module,expected] of [['prices','Price A'],['accounting','Work A']]){
      const payload={tenantId:firstCompany.tenantId,kind,module,commessaId:job};
      const result=await viaHttp(owner.token,'listCompanyModule',payload);
      assert.equal(result.status,200);assert.deepEqual(result.result.rows.map(x=>x.name),[expected]);
      assert.equal((await viaHttp(second.token,'listCompanyModule',payload)).status,403);
      assert.equal((await viaHttp(owner.token,'listCompanyModule',{...payload,tenantId:secondCompany.tenantId})).status,403);
      assert.equal((await viaHttp(owner.token,'listCompanyModule',{...payload,commessaId:'missing-parent'})).status,400);
      assert.equal((await viaHttp(owner.token,'listCompanyModule',{...payload,commessaId:'../other'})).status,400);
    }
    const priceB=await viaHttp(second.token,'listCompanyModule',{tenantId:secondCompany.tenantId,kind,module:'prices',commessaId:job});
    assert.equal(priceB.result.rows[0].base,999);
  }
  assert.equal((await viaHttp(owner.token,'listCompanyModule',{tenantId:firstCompany.tenantId,kind:'ordinary',module:'members'})).status,400);
  assert.equal((await viaHttp(owner.token,'listCompanyModule',{tenantId:firstCompany.tenantId,kind:'ordinary',module:'__proto__'})).status,400);
  await call(admin,'setCompanyStatus',{tenantId:firstCompany.tenantId,status:'suspended',reason:'Module suspension test'});
  assert.equal((await viaHttp(owner.token,'listCompanyModule',{tenantId:firstCompany.tenantId,kind:'ordinary',module:'personnel'})).status,403);
  await call(admin,'setCompanyStatus',{tenantId:firstCompany.tenantId,status:'active',reason:'Module reactivation test'});
  console.log('Personnel, teams, vehicles, accounting and price-list tenant module assertions passed');
  // Support services: document ownership, personal event and notification privacy,
  // monthly aggregate, ordinary/snow hours and fail-closed integrations.
  const ownerId=owner.uid, otherId=second.uid;
  for(const [module,collection,field,data] of [
    ['documents','documents','ownerUserId',{name:'My document A',visibility:'personal',fileUrl:'https://example.invalid/SECRET',storagePath:'SECRET',sharedUserIds:['private']}],
    ['events','calendarEvents','createdByUid',{title:'My event A',startDate:'2026-11-26',type:'altro',description:'SECRET'}],
    ['notifications','appNotifications','createdByUid',{title:'My notification A',eventType:'test',body:'SECRET',createdByEmail:'private@example.invalid'}]
  ]){
    await adapter.db.doc(ar+'/'+collection+'/mine').set({...data,[field]:ownerId});
    await adapter.db.doc(ar+'/'+collection+'/other').set({...data,name:'Other private',title:'Other private',[field]:otherId});
    await adapter.db.doc(br+'/'+collection+'/mine').set({...data,[field]:ownerId,name:'Other tenant',title:'Other tenant'});
    await adapter.db.doc(collection+'/mine').set({...data,[field]:ownerId,name:'Global legacy',title:'Global legacy'});
    const payload={tenantId:firstCompany.tenantId,module,kind:'ordinary'};
    const result=await viaHttp(owner.token,'listCompanyModule',payload);
    assert.equal(result.status,200);assert.deepEqual(result.result.rows.map(x=>x.id),['mine']);
    assert.ok(!JSON.stringify(result.result).includes('SECRET'));assert.ok(!JSON.stringify(result.result).includes('private@example.invalid'));
    assert.equal((await viaHttp(second.token,'listCompanyModule',payload)).status,403);
    assert.equal((await viaHttp(admin.token,'listCompanyModule',payload)).status,403);
    const empty=await viaHttp(emptyOwner.token,'listCompanyModule',{...payload,tenantId:emptyCompany.tenantId});
    assert.equal(empty.status,200);assert.deepEqual(empty.result.rows,[]);
  }
  for(const [kind,hours,approvals] of [['ordinary','oreReports','oreApprovalRequests'],['snow','neve_ore','neve_ore_richieste']]){
    for(const [module,collection] of [['hours',hours],['approvals',approvals]]){
      await adapter.db.doc(ar+'/'+collection+'/same-id').set({date:'2026-11-26',commessaName:'Hours A',entries:[{operatore:'SECRET'}],status:'approved'});
      await adapter.db.doc(br+'/'+collection+'/same-id').set({date:'2026-11-26',commessaName:'Hours B',entries:[{}]});
      await adapter.db.doc(collection+'/same-id').set({commessaName:'Global legacy'});
      const payload={tenantId:firstCompany.tenantId,module,kind};
      const response=await viaHttp(owner.token,'listCompanyModule',payload);
      assert.equal(response.status,200);assert.deepEqual(response.result.rows.map(x=>x.name),['Hours A']);
      assert.equal(response.result.rows[0].count,1);assert.ok(!JSON.stringify(response.result).includes('SECRET'));
      assert.equal((await viaHttp(second.token,'listCompanyModule',payload)).status,403);
      assert.equal((await viaHttp(owner.token,'listCompanyModule',{...payload,tenantId:secondCompany.tenantId})).status,403);
    }
  }
  await adapter.db.doc(ar+'/sharedStaticViews/calendario__2026-11').set({payload:{month:'2026-11',reports:[{person:'SECRET'}],activities:[{name:'SECRET'},{name:'SECRET'}]}});
  await adapter.db.doc(br+'/sharedStaticViews/calendario__2026-11').set({payload:{month:'2026-11',reports:[{},{},{}],activities:[]}});
  const calendarPayload={tenantId:firstCompany.tenantId,module:'calendar',kind:'ordinary',month:'2026-11'};
  const calendar=await viaHttp(owner.token,'listCompanyModule',calendarPayload);
  assert.equal(calendar.status,200);assert.deepEqual(calendar.result.rows,[{id:'2026-11',name:'Riepilogo 2026-11',count:1,activitiesCount:2}]);
  assert.equal((await viaHttp(owner.token,'listCompanyModule',{...calendarPayload,month:'2026-13'})).status,400);
  assert.equal((await viaHttp(second.token,'listCompanyModule',calendarPayload)).status,403);
  const emptyCalendar=await viaHttp(emptyOwner.token,'listCompanyModule',{...calendarPayload,tenantId:emptyCompany.tenantId});
  assert.deepEqual(emptyCalendar.result.rows,[]);
  await adapter.db.doc('appConfig/driveAdminSecret').set({token:'SECRET'});
  await adapter.db.doc(ar+'/appConfig/driveBridge').set({token:'SECRET'});
  const integrationPayload={tenantId:firstCompany.tenantId,module:'integrations',kind:'ordinary'};
  const integrations=await viaHttp(owner.token,'listCompanyModule',integrationPayload);
  assert.equal(integrations.status,200);assert.equal(integrations.result.rows.length,5);
  assert.ok(integrations.result.rows.every(r=>r.status.startsWith('Non attiva')));
  assert.ok(!JSON.stringify(integrations.result).includes('SECRET'));
  assert.equal((await viaHttp(owner.token,'syncDrive',integrationPayload)).status,400);
  assert.equal((await viaHttp(second.token,'listCompanyModule',integrationPayload)).status,403);
  console.log('Documents, calendar, hours, notifications and fail-closed integration assertions passed');
  await adapter.auth.setCustomUserClaims(admin.uid,{});
  assert.equal((await viaHttp(admin.token,'listCompanies')).status,403);
  await rejects(()=>call(admin,'listCompanies'),'forbidden');
  await adapter.auth.updateUser(owner.uid,{disabled:true});
  await rejects(()=>call(owner,'invite',{tenantId,recipientEmail:worker.email,role:'worker'}),'unauthenticated');
  console.log('Firebase Auth emulator adapter integration passed');
}
main().catch(e=>{console.error(e.code||e.name);process.exitCode=1;}).finally(async()=>{if(http)await http.close();await adapter.close();});

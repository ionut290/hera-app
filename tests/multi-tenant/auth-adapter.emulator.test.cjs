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
  const fixtureLimits={jobs:100,plants:1000,admins:10,operators:100};
  // Explicit Super Admin authorization for existing synthetic regression fixtures.
  const call=(u,action,payload)=>adapter.dispatch(u.token,action,action==='createCompany'&&u.uid===admin.uid?{approvedLimits:fixtureLimits,...payload}:payload);
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
  // Writes through authenticated HTTP, with tenant-bound receipts and revision checks.
  const save=(u,p)=>viaHttp(u.token,'saveCompanyRecord',p);
  const baseWrite={tenantId:firstCompany.tenantId,kind:'ordinary',module:'jobs',requestId:randomUUID(),data:{nome:'New A',codice:'A01',parentCommessaId:null}};
  const createdJob=await save(owner,baseWrite);assert.equal(createdJob.status,200);
  const jobId=createdJob.result.recordId,jobPath=ar+'/commesse/'+jobId;
  assert.equal((await adapter.db.doc(jobPath).get()).data().nome,'New A');
  const replay=await save(owner,baseWrite);assert.equal(replay.status,200);assert.equal(replay.result.recordId,jobId);
  assert.equal((await save(owner,{...baseWrite,data:{...baseWrite.data,nome:'Different'}})).status,400);
  assert.equal((await adapter.db.doc(br+'/commesse/'+jobId).get()).exists,false);
  assert.equal((await save(second,baseWrite)).status,403);assert.equal((await save(admin,baseWrite)).status,403);
  const detail=await viaHttp(owner.token,'getCompanyRecord',{tenantId:firstCompany.tenantId,module:'jobs',kind:'ordinary',recordId:jobId});
  assert.equal(detail.status,200);assert.equal(detail.result.revision,1);
  const edit={...baseWrite,recordId:jobId,revision:1,requestId:randomUUID(),data:{...baseWrite.data,nome:'Updated A'}};
  assert.equal((await save(owner,edit)).status,200);
  assert.equal((await save(owner,{...edit,requestId:randomUUID()})).status,400);
  assert.equal((await adapter.db.doc(jobPath).get()).data().nome,'Updated A');
  assert.equal((await save(owner,{...edit,requestId:randomUUID(),revision:2,data:{...edit.data,tenantRevision:99}})).status,400);
  assert.equal((await save(owner,{...baseWrite,requestId:randomUUID(),recordId:'job-a',revision:0})).status,400);
  assert.equal((await save(owner,{...baseWrite,requestId:randomUUID(),data:{...baseWrite.data,parentCommessaId:'job-b'}})).status,400);
  const sub=await save(owner,{...baseWrite,requestId:randomUUID(),data:{nome:'Child A',codice:'CHILD',parentCommessaId:jobId}});assert.equal(sub.status,200);
  assert.equal((await adapter.db.doc(ar+'/commesse/'+sub.result.recordId).get()).data().parentCommessaId,jobId);
  const snowWrite=await save(owner,{...baseWrite,kind:'snow',requestId:randomUUID()});assert.equal(snowWrite.status,200);
  assert.equal((await adapter.db.doc(ar+'/neve_commesse/'+snowWrite.result.recordId).get()).exists,true);
  for(const [module,collection,data] of [['personnel','personale',{cognome:'Demo',nome:'Worker'}],['vehicles','mezzi',{nId:'DEMO-1',marca:'Demo',modello:'Vehicle'}]]){
    const payload={tenantId:firstCompany.tenantId,kind:'ordinary',module,requestId:randomUUID(),data};
    const saved=await save(owner,payload);assert.equal(saved.status,200);
    assert.equal((await adapter.db.doc(ar+'/'+collection+'/'+saved.result.recordId).get()).exists,true);
    assert.equal((await adapter.db.doc(br+'/'+collection+'/'+saved.result.recordId).get()).exists,false);
    const visible=await viaHttp(owner.token,'listCompanyModule',{tenantId:firstCompany.tenantId,module,kind:'ordinary'});
    assert.ok(visible.result.rows.some(r=>r.id===saved.result.recordId));
    assert.equal((await save(second,payload)).status,403);
    assert.equal((await save(owner,{...payload,recordId:saved.result.recordId,revision:1,requestId:randomUUID()})).status,200);
  }
  const pricePayload={tenantId:firstCompany.tenantId,kind:'ordinary',module:'prices',commessaId:jobId,requestId:randomUUID(),data:{codiceVoce:'PX',descrizione:'Synthetic price',unitaMisura:'h',prezzoBase:100,percentualeRibasso:0.2}};
  const newPrice=await save(owner,pricePayload);assert.equal(newPrice.status,200);
  const pricePath=jobPath+'/prezziario/'+newPrice.result.recordId;
  assert.equal((await adapter.db.doc(pricePath).get()).data().prezzoRibassato,80);
  assert.equal((await save(owner,{...pricePayload,requestId:randomUUID()})).status,400);
  assert.equal((await save(owner,{...pricePayload,requestId:randomUUID(),recordId:newPrice.result.recordId,revision:1,data:{...pricePayload.data,prezzoBase:200}})).status,200);
  assert.equal((await adapter.db.doc(pricePath).get()).data().prezzoRibassato,160);
  await adapter.db.doc(jobPath+'/lavorazioni/linked').set({codiceVocePrezzo:'PX',totale:160});
  assert.equal((await save(owner,{...pricePayload,requestId:randomUUID(),recordId:newPrice.result.recordId,revision:2})).status,400);
  assert.equal((await adapter.db.doc(jobPath+'/lavorazioni/linked').get()).data().totale,160);
  assert.equal((await save(owner,{...baseWrite,requestId:randomUUID(),module:'impianti'})).status,400);
  await call(admin,'setCompanyStatus',{tenantId:firstCompany.tenantId,status:'suspended',reason:'Write suspension test'});
  assert.equal((await save(owner,{...baseWrite,requestId:randomUUID()})).status,403);
  await call(admin,'setCompanyStatus',{tenantId:firstCompany.tenantId,status:'active',reason:'Write reactivation test'});
  console.log('Transactional laboratory job, registry and unlinked price writes passed');
  // Operational pilot on a fresh synthetic company, not the existing seeded hours data.
  const opsOwner=await account('ops-owner'),opsB=await account('ops-other');
  const oa=await call(admin,'createCompany',{name:'Operations A',ownerEmail:opsOwner.email}),ob=await call(admin,'createCompany',{name:'Operations B',ownerEmail:opsB.email});
  const opRoot='tenants/'+oa.tenantId,opOther='tenants/'+ob.tenantId;
  const createBase=async(module,data,kind='ordinary')=>{const res=await save(opsOwner,{tenantId:oa.tenantId,module,kind,requestId:randomUUID(),data});assert.equal(res.status,200);return res.result.recordId;};
  const j=await createBase('jobs',{nome:'Operations job'}),j2=await createBase('jobs',{nome:'Second job'});
  const leader=await createBase('personnel',{cognome:'Demo',nome:'Leader'}),person=await createBase('personnel',{cognome:'Demo',nome:'Worker'}),vehicle=await createBase('vehicles',{nId:'DEMO-V'});
  const op=(u,operation,data,extra={})=>viaHttp(u.token,'createCompanyOperation',{tenantId:oa.tenantId,commessaId:j,kind:'ordinary',operation,requestId:randomUUID(),data,...extra});
  const p={denominazione:'Synthetic plant',comune:'Demo',gpsY:44.5,gpsX:11.3};
  const requestId=randomUUID(),plant=await op(opsOwner,'plant',p,{requestId});assert.equal(plant.status,200);
  const plantId=plant.result.recordId;
  assert.equal((await op(opsOwner,'plant',p,{requestId})).result.recordId,plantId);
  assert.equal((await op(opsOwner,'plant',{...p,done:true})).status,400);
  assert.equal((await op(opsOwner,'plant',p)).status,400);
  assert.equal((await op(opsB,'plant',p)).status,403);
  assert.equal((await adapter.db.doc(opRoot+'/commesse/'+j+'/impianti/'+plantId).get()).data().done,false);
  assert.equal((await adapter.db.doc(opRoot+'/commesse/'+j+'/impiantiFisici/'+plantId).get()).exists,true);
  assert.equal((await adapter.db.doc(opOther+'/commesse/'+j+'/impianti/'+plantId).get()).exists,false);
  const preview=await viaHttp(opsOwner.token,'listCompanyModule',{tenantId:oa.tenantId,commessaId:j,kind:'ordinary',module:'plants'});
  assert.equal(preview.status,200);assert.deepEqual(preview.result.rows.map(r=>r.name),['Synthetic plant']);
  const team={date:'2026-11-26',leaderId:leader,personnelIds:[person],vehicleIds:[vehicle],plantIds:[plantId]};
  assert.equal((await op(opsOwner,'team',{...team,leaderId:'foreign'})).status,400);
  const assigned=await op(opsOwner,'team',team);assert.equal(assigned.status,200);
  assert.equal((await adapter.db.doc(opRoot+'/squadreCommesse/'+j).get()).data().squadre[0].caposquadra,'Demo Leader');
  assert.equal((await adapter.db.doc(opRoot+'/squadreStorico/2026-11-26__'+j).get()).exists,true);
  assert.equal((await op(opsOwner,'team',{...team,plantIds:[]},{commessaId:j2})).status,400);
  assert.equal((await op(opsB,'team',team)).status,403);
  const hours={date:'2026-11-26',personnelId:person,hours:7.5};
  const hourId=randomUUID(),pending=await op(opsOwner,'hoursRequest',hours,{requestId:hourId});assert.equal(pending.status,200);
  assert.equal((await op(opsOwner,'hoursRequest',hours,{requestId:hourId})).result.recordId,pending.result.recordId);
  assert.equal((await adapter.db.doc(opRoot+'/oreApprovalRequests/'+pending.result.recordId).get()).data().status,'pending_level1');
  assert.equal((await adapter.db.collection(opRoot+'/oreReports').get()).empty,true);
  assert.equal((await op(opsOwner,'hoursRequest',hours)).status,400);
  assert.equal((await op(opsOwner,'hoursRequest',{...hours,hours:25})).status,400);
  assert.equal((await op(opsB,'hoursRequest',hours)).status,403);
  const sj=await createBase('jobs',{nome:'Snow operations'},'snow');
  const sp=await op(opsOwner,'plant',{denominazione:'Synthetic snow plant'},{kind:'snow',commessaId:sj});assert.equal(sp.status,200);
  assert.equal((await adapter.db.doc(opRoot+'/neve_commesse/'+sj+'/impianti/'+sp.result.recordId).get()).exists,true);
  await call(admin,'setCompanyStatus',{tenantId:oa.tenantId,status:'suspended',reason:'Operational suspension test'});
  assert.equal((await op(opsOwner,'plant',{denominazione:'Blocked'})).status,403);
  console.log('Laboratory operational plant, team assignment and pending hours assertions passed');
  // Self-service starter plan; all counters are enforced in the same transactions.
  const manualOwner=await account('manual-starter-owner');
  const manual=await viaHttp(admin.token,'createCompany',{name:'Manual starter',ownerEmail:manualOwner.email});
  assert.equal(manual.status,200);
  assert.deepEqual((await viaHttp(manualOwner.token,'getCompanyPlan',{tenantId:manual.result.tenantId})).result.plan.limits,{jobs:1,plants:10,admins:1,operators:3});
  const starterOwner=await account('starter-owner');
  const self=await viaHttp(starterOwner.token,'registerCompany',{name:'Starter company'});
  assert.equal(self.status,200);const st=self.result.tenantId;
  assert.equal((await viaHttp(starterOwner.token,'registerCompany',{name:'Second attempt'})).result.tenantId,st);
  assert.equal((await viaHttp(starterOwner.token,'registerCompany',{name:'Override',limits:{jobs:100}})).status,400);
  const plan=await viaHttp(starterOwner.token,'getCompanyPlan',{tenantId:st});
  assert.deepEqual(plan.result.plan.limits,{jobs:1,plants:10,admins:1,operators:3});
  assert.equal(plan.result.plan.usage.admins,1);
  assert.equal((await viaHttp(worker.token,'getCompanyPlan',{tenantId:st})).status,403);
  const sjob={tenantId:st,module:'jobs',kind:'ordinary',requestId:randomUUID(),data:{nome:'Starter job'}};
  const createdStarterJob=await save(starterOwner,sjob);assert.equal(createdStarterJob.status,200);
  assert.equal((await save(starterOwner,sjob)).result.recordId,createdStarterJob.result.recordId);
  assert.equal((await save(starterOwner,{...sjob,kind:'snow',requestId:randomUUID()})).code,'invalid-quota-jobs');
  for(let i=0;i<11;i++){
    const r=await viaHttp(starterOwner.token,'createCompanyOperation',{tenantId:st,operation:'plant',kind:'ordinary',commessaId:createdStarterJob.result.recordId,requestId:randomUUID(),data:{denominazione:'Starter plant '+i}});
    if(i<10)assert.equal(r.status,200);else assert.equal(r.code,'invalid-quota-plants');
  }
  await rejects(()=>call(starterOwner,'invite',{tenantId:st,recipientEmail:worker.email,role:'admin'}),'invalid-quota-admins');
  const invitees=await Promise.all([0,1,2,3].map(i=>account('starter-worker-'+i)));
  // Four invites may exist; only three concurrent acceptances can consume a seat.
  const invitations=await Promise.all(invitees.map(u=>call(starterOwner,'invite',{tenantId:st,recipientEmail:u.email,role:'worker'})));
  const accepted=await Promise.allSettled(invitees.map((u,i)=>call(u,'accept',{tenantId:st,inviteToken:invitations[i].token})));
  assert.equal(accepted.filter(x=>x.status==='fulfilled').length,3);
  assert.equal(accepted.filter(x=>x.status==='rejected'&&x.reason.code==='invalid-quota-operators').length,1);
  for(let i=0;i<4;i++){const r=await save(starterOwner,{tenantId:st,module:'personnel',kind:i%2?'snow':'ordinary',requestId:randomUUID(),data:{nome:'Person '+i,cognome:'Starter'}});if(i<3)assert.equal(r.status,200);else assert.equal(r.code,'invalid-quota-personnel');}
  const limits={jobs:2,plants:20,admins:2,operators:5};
  const req=await viaHttp(starterOwner.token,'requestPlanUpgrade',{tenantId:st,limits,reason:'Need more capacity'});
  assert.equal(req.status,200);
  assert.equal((await viaHttp(starterOwner.token,'reviewPlanUpgrade',{tenantId:st,requestId:req.result.requestId,decision:'approve',reason:'Self approval'})).status,403);
  assert.equal((await viaHttp(admin.token,'getPlanForAdmin',{tenantId:st})).result.plan.pending.id,req.result.requestId);
  assert.equal((await viaHttp(starterOwner.token,'listPendingPlanUpgrades',{})).status,403);
  const inbox=await viaHttp(admin.token,'listPendingPlanUpgrades',{limit:25});assert.equal(inbox.status,200);assert.ok(inbox.result.requests.some(r=>r.tenantId===st&&r.requestId===req.result.requestId));
  assert.equal((await save(starterOwner,{...sjob,requestId:randomUUID()})).code,'invalid-quota-jobs');
  assert.equal((await viaHttp(admin.token,'reviewPlanUpgrade',{tenantId:st,requestId:req.result.requestId,decision:'approve',limits:{...limits,plants:25},reason:'Approved synthetic custom limits'})).status,200);
  assert.equal((await save(starterOwner,{...sjob,kind:'snow',requestId:randomUUID()})).status,200);
  assert.deepEqual((await viaHttp(starterOwner.token,'getCompanyPlan',{tenantId:st})).result.plan.limits,{...limits,plants:25});
  const req2=await viaHttp(starterOwner.token,'requestPlanUpgrade',{tenantId:st,limits:{...limits,plants:25,jobs:3},reason:'More jobs'});
  assert.equal((await viaHttp(admin.token,'reviewPlanUpgrade',{tenantId:st,requestId:req2.result.requestId,decision:'reject',reason:'Not approved'})).status,200);
  assert.deepEqual((await viaHttp(starterOwner.token,'getCompanyPlan',{tenantId:st})).result.plan.limits,{...limits,plants:25});
  assert.equal((await viaHttp(admin.token,'listPendingPlanUpgrades',{})).result.requests.some(r=>r.tenantId===st),false);
  assert.equal((await viaHttp(starterOwner.token,'getCompanyPlan',{tenantId:st})).result.plan.lastDecision.decision,'reject');
  console.log('Starter self-registration, transactional quotas, manual defaults, custom approvals and super-admin inbox passed');
  // A verified account without a historical platformUsers profile can register
  // or accept a tenant invitation through the isolated service without Super Admin.
  const freshOwner=await account('autonomous-no-profile'),freshWorker=await account('autonomous-invitee'),otherWorker=await account('autonomous-wrong-recipient');
  assert.equal((await adapter.db.doc('platformUsers/'+freshOwner.uid).get()).exists,false);
  const autonomous=await viaHttp(freshOwner.token,'registerCompany',{name:'Autonomous company'});
  assert.equal(autonomous.status,200);const autonomousId=autonomous.result.tenantId;
  assert.equal((await adapter.db.doc('platformUsers/'+freshOwner.uid).get()).exists,false);
  assert.equal((await viaHttp(freshOwner.token,'openCompany',{tenantId:autonomousId})).result.company.role,'owner');
  const generated=await viaHttp(freshOwner.token,'invite',{tenantId:autonomousId,recipientEmail:freshWorker.email,role:'worker'});
  assert.equal(generated.status,200);
  assert.equal((await viaHttp(otherWorker.token,'accept',{tenantId:autonomousId,inviteToken:generated.result.token})).code,'invalid-invite');
  assert.equal((await viaHttp(freshWorker.token,'accept',{tenantId:autonomousId,inviteToken:generated.result.token})).status,200);
  assert.equal((await viaHttp(freshWorker.token,'accept',{tenantId:autonomousId,inviteToken:generated.result.token})).code,'invalid-invite');
  assert.equal((await adapter.db.doc('platformUsers/'+freshWorker.uid).get()).exists,false);
  assert.equal((await viaHttp(freshWorker.token,'openCompany',{tenantId:autonomousId})).result.company.role,'worker');
  const people=await viaHttp(freshOwner.token,'listCompanyMembers',{tenantId:autonomousId,limit:25});
  assert.equal(people.status,200);assert.deepEqual(people.result.members.map(m=>m.role).sort(),['owner','worker']);
  assert.ok(!JSON.stringify(people.result).includes(generated.result.token));
  assert.equal((await viaHttp(freshWorker.token,'listCompanyMembers',{tenantId:autonomousId})).status,403);
  assert.equal((await viaHttp(otherWorker.token,'listCompanyMembers',{tenantId:autonomousId})).status,403);
  assert.equal((await viaHttp(admin.token,'listCompanyMembers',{tenantId:autonomousId})).status,403);
  assert.equal((await viaHttp(freshOwner.token,'getCompanyPlan',{tenantId:autonomousId})).result.plan.usage.operators,1);
  console.log('Verified-account self-service and invitation HTTP workflow passed without legacy profile');
  // Entry authorization must not create/approve legacy profiles.
  assert.equal((await viaHttp(freshOwner.token,'checkTenantEntry')).result.eligible,true);
  assert.equal((await viaHttp(unverified.token,'checkTenantEntry')).status,401);
  assert.equal((await viaHttp(review.token,'checkTenantEntry')).status,401);
  const blockedEntry=await account('blocked-legacy-entry');
  await adapter.db.doc('platformUsers/'+blockedEntry.uid).set({statoAccount:'bloccato',banned:true});
  assert.equal((await viaHttp(blockedEntry.token,'checkTenantEntry')).status,403);
  assert.equal((await viaHttp(blockedEntry.token,'registerCompany',{name:'Blocked account'})).status,403);
  assert.equal((await viaHttp(blockedEntry.token,'listMyCompanies')).status,403);
  assert.equal((await adapter.db.doc('platformUsers/'+freshOwner.uid).get()).exists,false);
  console.log('Tenant-only entry does not approve legacy profiles and rejects blocked accounts');
  await adapter.auth.setCustomUserClaims(admin.uid,{});
  assert.equal((await viaHttp(admin.token,'listCompanies')).status,403);
  await rejects(()=>call(admin,'listCompanies'),'forbidden');
  await adapter.auth.updateUser(owner.uid,{disabled:true});
  await rejects(()=>call(owner,'invite',{tenantId,recipientEmail:worker.email,role:'worker'}),'unauthenticated');
  console.log('Firebase Auth emulator adapter integration passed');
}
main().catch(e=>{console.error(e.code||e.name);process.exitCode=1;}).finally(async()=>{if(http)await http.close();await adapter.close();});

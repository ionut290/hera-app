'use strict';
// Actual original loader, read-only isolated Firestore paths, two synthetic sessions.
// Auth claims are synthetic rules fixtures, not a proof of real device token issuance.
const assert=require('node:assert/strict'),vm=require('node:vm');
const {readFileSync}=require('node:fs'),{join}=require('node:path');
const {initializeTestEnvironment,assertFails,assertSucceeds}=require('@firebase/rules-unit-testing');
const {doc,collection,setDoc,getDocFromServer,getDocsFromServer,query,where,orderBy,updateDoc,writeBatch}=require('firebase/firestore');
const {createFlatSessionScope}=require('./flat-session-scope.cjs');
const source=readFileSync(join(__dirname,'../../app.js'),'utf8');
function section(start,end){const a=source.indexOf(start),b=source.indexOf(end,a+start.length);assert.ok(a>=0&&b>a,'Original function changed: '+start);return source.slice(a,b).trim();}
const original=[
 section('function normalizeCommessaDocument(doc) {','function normalizeSquadraStoricoDocument('),
 section('function subscribeCommesse() {','function stopCommesseSubscription() {'),
 section('function getCommesseCollectionName() {','function getSquadreCurrentCollectionName() {')
].join('\n');
function runtime(scope,uid){
 const c=vm.createContext({db:scope.db,currentUser:{uid},snow:false,console:{log(){},warn(){},error(){}},commesseById:new Map(),commesseLoadState:{status:'idle'},unsubscribeCommesse:null,renderCommesseHomeList(){},startCommesseLoadTimeout(){},clearCommesseLoadTimeout(){},validateFirebaseConfigForCommesse:()=>true,refreshCommesseDependentUI(){},tryAutoOpenAssignedCommessaAtStartup(){},renderNextActionCard(){},renderTodaySummary(){},logFirestoreError(){},getReadableFirestoreError:()=> 'Denied',runFirestoreGetWithRetry:q=>q.get(),runAfterFirstRender:fn=>fn(),isSnowServiceContext:()=>c.snow});
 vm.runInContext('function stopCommesseSubscription(){if(unsubscribeCommesse)unsubscribeCommesse();unsubscribeCommesse=null;} function saveCommesseLocalCache(){} function loadCommesseFromLocalCache(){commesseById=new Map();return [];}\n'+original,c,{filename:'original-app-commesse.js'});
 return {c,load:()=>vm.runInContext('subscribeCommesse()',c),names:()=>JSON.parse(vm.runInContext('JSON.stringify(Array.from(commesseById.values(),d=>d.nome))',c)),status:()=>c.commesseLoadState.status,close(){vm.runInContext('stopCommesseSubscription();commesseById=new Map();currentUser=null;',c);scope.close();}};
}
async function main(){
 assert.equal(process.env.FIRESTORE_EMULATOR_HOST,'127.0.0.1:8180');
 assert.equal(process.env.FIREBASE_AUTH_EMULATOR_HOST,'127.0.0.1:9199');
 const env=await initializeTestEnvironment({projectId:'demo-varga-multi-tenant',firestore:{host:'127.0.0.1',port:8180,rules:readFileSync(join(__dirname,'flat-session-experiment.rules'),'utf8')}});
 const uid='flat-session-owner',a='session-a',b='session-b',areas=['commesse','neve_commesse'],children=['impianti','impiantiFisici','lavorazioni','prezziario'];
 const job=(area,tenant)=>'session-'+area+'-'+tenant;
 const sessionPath=(u,id)=>'experimentSessions/'+u+'/devices/'+id;
 const session=(tenantId,status='active',expired=false)=>({tenantId,status,expiresAt:new Date(Date.now()+(expired?-60000:3600000))});
 const seed=(p,d)=>env.withSecurityRulesDisabled(ctx=>setDoc(doc(ctx.firestore(),p),d));
 const claims=id=>({email_verified:true,vargaExperimentSession:id});
 const runtimes=[];
 try{
  await env.withSecurityRulesDisabled(async ctx=>{
   const db=ctx.firestore(),put=(p,d)=>setDoc(doc(db,p),d);
   for(const tenant of [a,b]){
    await put('tenants/'+tenant,{status:'active'});
    await put('tenants/'+tenant+'/members/'+uid,{status:'active',role:'owner'});
    for(const area of areas){
     const path=area+'/'+job(area,tenant);
     await put(path,{tenantId:tenant,flatSessionExperiment:true,nome:tenant+' '+area,createdAt:new Date()});
     for(const child of children)await put(path+'/'+child+'/same-id',{marker:tenant,stato:'DA FARE'});
    }
   }
   await put('commesse/session-legacy-global',{nome:'Global legacy',createdAt:new Date()});
   await put(sessionPath(uid,'phone'),session(a));
   await put(sessionPath(uid,'tablet'),session(b));
   await put(sessionPath(uid,'expired'),session(a,'active',true));
   await put(sessionPath(uid,'revoked'),session(a,'revoked'));
  });
  const client=(u,c)=>env.authenticatedContext(u,c).firestore();
  const phone=client(uid,claims('phone')),tablet=client(uid,claims('tablet'));
  const other=client('different-user',claims('phone'));
  const bad=[client(uid,{email_verified:true}),client(uid,claims('expired')),client(uid,claims('revoked')),client(uid,claims('absent')),client(uid,claims('bad/path')),client(uid,{...claims('phone'),apple_review:true}),client(uid,{...claims('phone'),email_verified:false}),other];
  const fq=(db,area,tenant)=>query(collection(db,area),where('tenantId','==',tenant),where('flatSessionExperiment','==',true),orderBy('createdAt','desc'));
  const mk=(db,tenantId,sessionId)=>{
   let current={uid,tenantId,sessionId};
   const scope=createFlatSessionScope({firestore:db,tenantId,uid,sessionId,currentIdentity:()=>current});
   const r=runtime(scope,uid);runtimes.push(r);return {r,scope,change(v){current=v;}};
  };
  const x=mk(phone,a,'phone'),y=mk(tablet,b,'tablet');
  await Promise.all([x.r.load(),y.r.load()]);
  assert.deepEqual(x.r.names(),[a+' commesse']);assert.deepEqual(y.r.names(),[b+' commesse']);
  x.r.c.snow=true;y.r.c.snow=true;
  await Promise.all([x.r.load(),y.r.load()]);
  assert.deepEqual(x.r.names(),[a+' neve_commesse']);assert.deepEqual(y.r.names(),[b+' neve_commesse']);
  for(const area of areas){
   await assertFails(getDocsFromServer(query(collection(phone,area),orderBy('createdAt','desc'))));
   assert.equal((await assertSucceeds(getDocsFromServer(fq(phone,area,a)))).size,1);
   assert.equal((await assertSucceeds(getDocsFromServer(fq(tablet,area,b)))).size,1);
   await assertFails(getDocsFromServer(fq(phone,area,b)));
   await assertFails(getDocsFromServer(fq(tablet,area,a)));
   for(const db of bad)await assertFails(getDocsFromServer(fq(db,area,a)));
   for(const child of children){
    const pa=area+'/'+job(area,a)+'/'+child+'/same-id',pb=area+'/'+job(area,b)+'/'+child+'/same-id';
    assert.equal((await assertSucceeds(getDocFromServer(doc(phone,pa)))).data().marker,a);
    assert.equal((await assertSucceeds(getDocFromServer(doc(tablet,pb)))).data().marker,b);
    await assertFails(getDocFromServer(doc(phone,pb)));
    await assertFails(getDocFromServer(doc(tablet,pa)));
    await assertFails(updateDoc(doc(phone,pa),{stato:'FATTO'}));
   }
   const batch=writeBatch(phone);batch.update(doc(phone,area+'/'+job(area,a)+'/impianti/same-id'),{stato:'FATTO'});batch.update(doc(phone,area+'/'+job(area,b)+'/impianti/same-id'),{stato:'FATTO'});await assertFails(batch.commit());
   assert.equal((await getDocFromServer(doc(phone,area+'/'+job(area,a)+'/impianti/same-id'))).data().stato,'DA FARE');
  }
  await assertFails(getDocFromServer(doc(phone,'commesse/session-legacy-global')));
  await assertFails(setDoc(doc(phone,sessionPath(uid,'phone')),session(b)));
  await assertFails(setDoc(doc(phone,sessionPath(uid,'new')),session(b)));
  await assertSucceeds(getDocFromServer(doc(phone,sessionPath(uid,'phone'))));
  await assertFails(getDocFromServer(doc(phone,sessionPath(uid,'tablet'))));
  await assertFails(getDocsFromServer(collection(phone,'experimentSessions/'+uid+'/devices')));
  assert.throws(()=>x.scope.db.collection('platformUsers'),{code:'permission-denied'});
  assert.throws(()=>x.scope.db.collection('tenants/'+b+'/commesse'),{code:'permission-denied'});
  assert.equal(x.scope.db.collection('commesse').add,undefined);
  const old=x.scope.db.collection('commesse').orderBy('createdAt','desc');
  // A device's selected company changes without changing the other device.
  await seed(sessionPath(uid,'phone'),session(b));
  await assertFails(old.get());
  await assertSucceeds(getDocsFromServer(fq(tablet,'commesse',b)));
  x.change({uid,tenantId:b,sessionId:'phone'});
  await assert.rejects(old.get(),{code:'permission-denied'});x.r.close();
  const z=mk(phone,b,'phone');await z.r.load();assert.deepEqual(z.r.names(),[b+' commesse']);
  await seed(sessionPath(uid,'phone'),session(b,'revoked'));
  // A reused active listener can return its last server snapshot while a
  // dependent-rule change is propagating. Diagnose this separately from a
  // genuinely new server request; do not interpret a cached answer as denial.
  const beforeClose=await Promise.allSettled([getDocsFromServer(fq(phone,'commesse',b))]);
  console.log('Revoked active-query observation: '+beforeClose[0].status);
  z.r.close();
  const freshPhone=client(uid,claims('phone'));
  await assertFails(getDocFromServer(doc(freshPhone,'commesse/'+job('commesse',b))));
  await assertFails(getDocsFromServer(fq(freshPhone,'commesse',b)));
  await assertSucceeds(getDocsFromServer(fq(tablet,'commesse',b)));
  y.r.close();
  await seed('tenants/'+b+'/members/'+uid,{status:'revoked',role:'owner'});
  const freshTablet=client(uid,claims('tablet'));
  await assertFails(getDocFromServer(doc(freshTablet,'commesse/'+job('commesse',b))));
  await assertFails(getDocsFromServer(fq(freshTablet,'commesse',b)));
  console.log('Original commesse loader + flat per-device session rules: scoped normal/snow, cross-device isolation, denied writes, expiry and revocation passed.');
 }finally{for(const r of runtimes)r.close();await env.cleanup();}
}
main().catch(e=>{console.error(e);process.exitCode=1;});

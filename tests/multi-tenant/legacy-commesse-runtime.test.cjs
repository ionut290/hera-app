'use strict';
// Execute original app.js functions; UI/session/cache hooks are synthetic fixtures.
const {readFileSync}=require('node:fs'),{join}=require('node:path'),vm=require('node:vm'),assert=require('node:assert/strict');
const {initializeTestEnvironment,assertFails}=require('@firebase/rules-unit-testing');
const {doc,setDoc,updateDoc,getDoc}=require('firebase/firestore');
const {createLegacyCommesseScope}=require('./legacy-commesse-scope.cjs');
const source=readFileSync(join(__dirname,'../../app.js'),'utf8');
function section(start,end){const a=source.indexOf(start),b=source.indexOf(end,a+start.length);assert.ok(a>=0&&b>a,'Original function changed: '+start);return source.slice(a,b).trim();}
const original=[
  section('function normalizeCommessaDocument(doc) {','function normalizeSquadraStoricoDocument('),
  section('function subscribeCommesse() {','function stopCommesseSubscription() {'),
  section('function getCommesseCollectionName() {','function getSquadreCurrentCollectionName() {')
].join('\n');
function runtime(scope,uid){
  const c=vm.createContext({db:scope.db,currentUser:{uid},snow:false,console:{log(){},warn(){},error(){}},commesseById:new Map(),commesseLoadState:{status:'idle'},unsubscribeCommesse:null,renderCommesseHomeList(){},startCommesseLoadTimeout(){},clearCommesseLoadTimeout(){},validateFirebaseConfigForCommesse:()=>true,refreshCommesseDependentUI(){},tryAutoOpenAssignedCommessaAtStartup(){},renderNextActionCard(){},renderTodaySummary(){},logFirestoreError(){},getReadableFirestoreError:()=> 'Denied',runFirestoreGetWithRetry:q=>q.get(),runAfterFirstRender:fn=>fn(),isSnowServiceContext:()=>c.snow});
  vm.runInContext(`
    function stopCommesseSubscription(){if(unsubscribeCommesse)unsubscribeCommesse();unsubscribeCommesse=null;}
    function saveCommesseLocalCache(){}
    function loadCommesseFromLocalCache(){commesseById=new Map();return [];}
  `+original,c,{filename:'original-app-commesse.js'});
  return {c,load:()=>vm.runInContext('subscribeCommesse()',c),names:()=>JSON.parse(vm.runInContext('JSON.stringify(Array.from(commesseById.values(),d=>d.nome))',c)),status:()=>c.commesseLoadState.status,close:()=>{vm.runInContext('stopCommesseSubscription();commesseById=new Map();currentUser=null;',c);scope.close();}};
}
async function main(){
  assert.equal(process.env.FIRESTORE_EMULATOR_HOST,'127.0.0.1:8180');
  const env=await initializeTestEnvironment({projectId:'demo-varga-multi-tenant',firestore:{host:'127.0.0.1',port:8180,rules:readFileSync(join(__dirname,'firestore.rules'),'utf8')}});
  const scopes=[];
  try{
    await env.withSecurityRulesDisabled(async ctx=>{
      const db=ctx.firestore(),seed=(p,d)=>setDoc(doc(db,p),d),createdAt=new Date('2026-11-26T00:00:00Z');
      await Promise.all([
        seed('tenants/runtime-a',{status:'active'}),seed('tenants/runtime-b',{status:'active'}),
        seed('tenants/runtime-a/members/runtime-owner-a',{status:'active',role:'owner'}),
        seed('tenants/runtime-b/members/runtime-owner-b',{status:'active',role:'owner'}),
        seed('tenants/runtime-a/members/runtime-worker',{status:'active',role:'worker'}),
        seed('tenants/runtime-a/commesse/same-id',{nome:'Original A',createdAt}),
        seed('tenants/runtime-b/commesse/same-id',{nome:'Original B',createdAt}),
        seed('tenants/runtime-a/neve_commesse/same-id',{nome:'Snow A',createdAt}),
        seed('tenants/runtime-b/neve_commesse/same-id',{nome:'Snow B',createdAt}),
        seed('commesse/runtime-global',{nome:'Global legacy',createdAt}),seed('neve_commesse/runtime-global',{nome:'Global snow',createdAt})
      ]);
    });
    const fa=env.authenticatedContext('runtime-owner-a',{email_verified:true}).firestore(),fb=env.authenticatedContext('runtime-owner-b',{email_verified:true}).firestore();
    let identity='runtime-owner-a';
    const make=(firestore,tenantId,uid,currentUid)=>{const s=createLegacyCommesseScope({firestore,tenantId,uid,currentUid});scopes.push(s);return s;};
    const a=make(fa,'runtime-a','runtime-owner-a',()=>identity);
    assert.throws(()=>a.db.collection('platformUsers'),{code:'permission-denied'});
    assert.throws(()=>a.db.collection('tenants/runtime-b/commesse'),{code:'permission-denied'});
    assert.equal(a.db.collection('commesse').add,undefined);assert.equal(a.db.collection('commesse').doc,undefined);
    const ra=runtime(a,'runtime-owner-a');await ra.load();assert.deepEqual(ra.names(),['Original A']);assert.equal(ra.status(),'loaded');
    ra.c.snow=true;await ra.load();assert.deepEqual(ra.names(),['Snow A']);
    const old=a.db.collection('commesse').orderBy('createdAt','desc');ra.close();assert.deepEqual(ra.names(),[]);await assert.rejects(old.get(),{code:'permission-denied'});
    const rb=runtime(make(fb,'runtime-b','runtime-owner-b',()=> 'runtime-owner-b'),'runtime-owner-b');
    await rb.load();assert.deepEqual(rb.names(),['Original B']);rb.c.snow=true;await rb.load();assert.deepEqual(rb.names(),['Snow B']);rb.close();
    const rc=runtime(make(fa,'runtime-b','runtime-owner-a',()=>identity),'runtime-owner-a');
    await rc.load();assert.equal(rc.status(),'error');assert.deepEqual(rc.names(),[]);rc.close();
    const fw=env.authenticatedContext('runtime-worker',{email_verified:true}).firestore(),rw=runtime(make(fw,'runtime-a','runtime-worker',()=> 'runtime-worker'),'runtime-worker');
    await rw.load();assert.equal(rw.status(),'error');assert.deepEqual(rw.names(),[]);rw.close();
    await assertFails(getDoc(doc(fa,'commesse/runtime-global')));await assertFails(getDoc(doc(fa,'neve_commesse/runtime-global')));
    await assertFails(setDoc(doc(fa,'tenants/runtime-a/neve_commesse/forbidden'),{nome:'Forbidden'}));
    const rev=make(fa,'runtime-a','runtime-owner-a',()=>identity),rr=runtime(rev,'runtime-owner-a');
    await rr.load();assert.deepEqual(rr.names(),['Original A']);
    await env.withSecurityRulesDisabled(ctx=>updateDoc(doc(ctx.firestore(),'tenants/runtime-a/members/runtime-owner-a'),{status:'revoked'}));
    await rr.load();assert.equal(rr.status(),'error');assert.deepEqual(rr.names(),[]);
    identity='another-user';assert.throws(()=>rev.db.collection('commesse'),{code:'permission-denied'});rr.close();
    console.log('Original commesse loader: ordinary/snow isolation, disposal and revocation passed.');
  }finally{for(const s of scopes)s.close();await env.cleanup();}
}
main().catch(e=>{console.error(e);process.exitCode=1;});

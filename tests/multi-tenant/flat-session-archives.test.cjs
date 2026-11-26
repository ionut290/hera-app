'use strict';
// Read-only original archive selectors + real emulator rules and control-record monitor.
const assert=require('node:assert/strict'),vm=require('node:vm');
const {readFileSync}=require('node:fs'),{join}=require('node:path');
const {initializeTestEnvironment,assertFails,assertSucceeds}=require('@firebase/rules-unit-testing');
const {doc,collection,setDoc,getDocFromServer,getDocsFromServer,query,where,updateDoc}=require('firebase/firestore');
const {createSessionMonitor}=require('./flat-session-monitor.cjs');
const pairs={
 getCommesseCollectionName:['commesse','neve_commesse'],
 getPersonaleCollectionName:['personale','neve_personale'],
 getMezziCollectionName:['mezzi','neve_mezzi'],
 getSquadreCurrentCollectionName:['squadreCommesse','neve_squadre'],
 getSquadreHistoryCollectionName:['squadreStorico','neve_squadre_storico'],
 getOreReportsCollectionName:['oreReports','neve_ore'],
 getOreApprovalRequestsCollectionName:['oreApprovalRequests','neve_ore_richieste']
};
const source=readFileSync(join(__dirname,'../../app.js'),'utf8');
function extract(name){
 const start=source.indexOf('function '+name+'() {'),end=source.indexOf('\n}',start);
 assert(start>=0&&end>start,'Original selector not found: '+name);
 return source.slice(start,end+2);
}
const scope=vm.createContext({snow:false,isSnowServiceContext:()=>scope.snow});
vm.runInContext(Object.keys(pairs).map(extract).join('\n'),scope);
for(const [name,names] of Object.entries(pairs))for(let i=0;i<2;i++){scope.snow=Boolean(i);assert.equal(vm.runInContext(name+'()',scope),names[i]);}
async function main(){
 assert.equal(process.env.FIRESTORE_EMULATOR_HOST,'127.0.0.1:8180');
 assert.equal(process.env.FIREBASE_AUTH_EMULATOR_HOST,'127.0.0.1:9199');
 const env=await initializeTestEnvironment({projectId:'demo-varga-multi-tenant',firestore:{host:'127.0.0.1',port:8180,rules:readFileSync(join(__dirname,'flat-session-experiment.rules'),'utf8')}});
 const uid='archive-owner',a='archive-a',b='archive-b',sid='archive-phone',otherSid='archive-tablet',paths=Object.values(pairs).flat();
 const sp=(id)=>'experimentSessions/'+uid+'/devices/'+id;
 const record=(tenantId,status='active',expired=false)=>({tenantId,status,expiresAt:new Date(Date.now()+(expired?-1000:600000))});
 const seed=(p,data)=>env.withSecurityRulesDisabled(ctx=>setDoc(doc(ctx.firestore(),p),data));
 const client=(id,sessionId)=>env.authenticatedContext(id,{email_verified:true,vargaExperimentSession:sessionId}).firestore();
 const monitors=[];
 try{
  await env.withSecurityRulesDisabled(async ctx=>{
   const db=ctx.firestore(),put=(p,d)=>setDoc(doc(db,p),d);
   for(const tenant of [a,b]){
    await put('tenants/'+tenant,{status:'active'});
    await put('tenants/'+tenant+'/members/'+uid,{status:'active',role:'owner'});
    await put('tenants/'+tenant+'/members/archive-worker',{status:'active',role:'worker'});
    for(const archive of paths)await put(archive+'/'+tenant+'-record',{tenantId:tenant,flatSessionExperiment:true,nome:tenant+' '+archive,createdAt:new Date()});
   }
   for(const archive of paths)await put(archive+'/archive-legacy',{nome:'legacy private'});
   await put(sp(sid),record(a));await put(sp(otherSid),record(b));
   await put('experimentSessions/archive-worker/devices/worker-phone',record(a));
   await put(sp('expired'),record(a,'active',true));
  });
  const phone=client(uid,sid),tablet=client(uid,otherSid),worker=client('archive-worker','worker-phone');
  const q=(db,name,tenantId)=>query(collection(db,name),where('tenantId','==',tenantId),where('flatSessionExperiment','==',true));
  for(const name of paths){
   const pa=name+'/'+a+'-record',pb=name+'/'+b+'-record';
   assert.equal((await assertSucceeds(getDocsFromServer(q(phone,name,a)))).size,1);
   assert.equal((await assertSucceeds(getDocsFromServer(q(tablet,name,b)))).size,1);
   await assertFails(getDocsFromServer(q(phone,name,b)));
   await assertFails(getDocsFromServer(collection(phone,name)));
   await assertFails(getDocsFromServer(q(worker,name,a)));
   await assertFails(getDocFromServer(doc(phone,pb)));
   await assertFails(getDocFromServer(doc(phone,name+'/archive-legacy')));
   await assertFails(updateDoc(doc(phone,pa),{nome:'forbidden write'}));
  }
  const visible={a:['synthetic A'],b:['synthetic B']};
  const closed=[];
  let resolveClosed;const closedSignal=new Promise(resolve=>{resolveClosed=resolve;});
  const ma=createSessionMonitor({firestore:phone,uid,sessionId:sid,tenantId:a,pollMs:250,onClose:reason=>{visible.a=[];closed.push(['a',reason]);resolveClosed();}});
  const mb=createSessionMonitor({firestore:tablet,uid,sessionId:otherSid,tenantId:b,pollMs:250,onClose:reason=>{visible.b=[];closed.push(['b',reason]);}});
  monitors.push(ma,mb);
  assert.equal(await ma.start(),true);assert.equal(await mb.start(),true);ma.assertActive();mb.assertActive();
  await assertFails(getDocFromServer(doc(phone,sp(otherSid))));
  await assertFails(getDocsFromServer(collection(phone,'experimentSessions/'+uid+'/devices')));
  await assertFails(updateDoc(doc(phone,sp(sid)),{tenantId:b}));
  const expired=createSessionMonitor({firestore:client(uid,'expired'),uid,sessionId:'expired',tenantId:a,onClose:()=>{}});
  monitors.push(expired);assert.equal(await expired.start(),false);
  await seed(sp(sid),record(a,'revoked'));
  // Bounded lab observation; this does not assert instantaneous revocation.
  let timer;try{await Promise.race([closedSignal,new Promise((_,reject)=>{timer=setTimeout(()=>reject(Error('Monitor did not clear revoked session')),8000);})]);}finally{clearTimeout(timer);}
  assert.deepEqual(visible.a,[]);assert.deepEqual(visible.b,['synthetic B']);
  assert.throws(()=>ma.assertActive(),{code:'permission-denied'});mb.assertActive();
  assert.equal(closed.length,1);
  const fresh=client(uid,sid);
  await assertFails(getDocFromServer(doc(fresh,'personale/'+a+'-record')));
  await assertFails(getDocsFromServer(q(fresh,'oreReports',a)));
  await assertSucceeds(getDocsFromServer(q(tablet,'personale',b)));
  console.log('Original archive selectors: 14 roots, ordinary/snow isolation, own-session monitoring and bounded synthetic cache clearing passed.');
 }finally{for(const m of monitors)m.close();await env.cleanup();}
}
main().catch(e=>{console.error(e);process.exitCode=1;});

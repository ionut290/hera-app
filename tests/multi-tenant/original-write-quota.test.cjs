'use strict';
// Execute original app.js create handlers with synthetic UI and real server quota transactions.
const assert=require('node:assert/strict'),vm=require('node:vm');
const {randomUUID}=require('node:crypto'),{readFileSync}=require('node:fs'),{join}=require('node:path');
const {createLocalAdapter}=require('./auth-adapter.emulator.cjs');
const {createOriginalWriteScope}=require('./original-write-scope.cjs');
const {originalFunctions}=require('./original-functions.cjs');
const source=readFileSync(join(__dirname,'../../app.js'),'utf8'),extract=originalFunctions(source);
const original=['createCommessa','addManagementImpianto','validateManagementPlant'].map(extract).join('\n');
async function main(){
 assert.equal(process.env.FIRESTORE_EMULATOR_HOST,'127.0.0.1:8180');
 assert.equal(process.env.FIREBASE_AUTH_EMULATOR_HOST,'127.0.0.1:9199');
 const adapter=createLocalAdapter(),db=adapter.db,scopes=[];
 async function account(){
  const email=randomUUID()+'@example.invalid',password='Demo-'+randomUUID();
  const u=await adapter.auth.createUser({email,password,emailVerified:true});
  const response=await fetch('http://127.0.0.1:9199/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=demo-key',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({email,password,returnSecureToken:true}),redirect:'error'});
  const data=await response.json();assert.ok(response.ok&&data.idToken&&data.localId===u.uid);
  return {uid:u.uid,email,token:data.idToken};
 }
 const call=(user,action,payload)=>adapter.dispatch(user.token,action,payload);
 function fixture(user,tenantId,snow=false,plantIndex=1){
  let current=user.uid;
  const scope=createOriginalWriteScope({tenantId,uid:user.uid,currentUid:()=>current,dispatch:adapter.dispatch,token:async()=>user.token});
  scopes.push(scope);
  const u={uid:user.uid,email:user.email},prompts=['Parco demo '+plantIndex,'SAP-'+plantIndex];
  const c=vm.createContext({auth:{currentUser:u},currentUser:u,ui:{commessaName:{value:'Commessa demo'},commessaCode:{value:'V1'},commessaType:{value:'main'},commessaParent:{value:''},commessaForm:{reset(){}}},alert(){},canManageData:()=>true,db:scope.db,getCommesseCollectionName:()=>snow?'neve_commesse':'commesse',firebase:{firestore:{FieldValue:{serverTimestamp:()=>null}}},driveAccessToken:null,closeNewCommessaDialog(){},updateCommessaParentField(){},window:{prompt:()=>prompts.shift()},managementCommessaId:'',commesseById:new Map(),getCommessaCachedImpianti:()=>[],normalizeManagementSearch:v=>String(v).trim().toLowerCase(),getOperatorDisplayName:()=>user.uid,managementEditingId:'',console});
  vm.runInContext(original,c,{filename:'unchanged-original-create.js'});
  return {c,scope,disconnect(){current=null;}};
 }
 const click={preventDefault(){}};
 const denied=(p,code)=>assert.rejects(p,{code});
 try{
  const owner=await account(),other=await account(),worker=await account();
  const {tenantId:a}=await call(owner,'registerCompany',{name:'Original quota A'});
  const {tenantId:b}=await call(other,'registerCompany',{name:'Original quota B'});
  const ar=db.doc('tenants/'+a),br=db.doc('tenants/'+b);
  const normal=fixture(owner,a),snow=fixture(owner,a,true);
  // Same quota across normal and snow; simultaneous original handlers.
  const results=await Promise.allSettled([normal.c.createCommessa(click),snow.c.createCommessa(click)]);
  assert.equal(results.filter(x=>x.status==='fulfilled').length,1);
  assert.equal(results.find(x=>x.status==='rejected')?.reason?.code,'invalid-quota-jobs');
  const [ordinary,winter]=await Promise.all([ar.collection('commesse').get(),ar.collection('neve_commesse').get()]);
  assert.equal(ordinary.size+winter.size,1);
  const isSnow=winter.size===1,job=(isSnow?winter:ordinary).docs[0];
  assert.equal(job.data().excelModelVersion,2);
  assert.equal((await ar.get()).data().starterPlan.usage.jobs,1);
  assert.equal((await br.collection('commesse').get()).size,0);
  assert.equal((await br.collection('neve_commesse').get()).size,0);
  const bFixture=fixture(other,b);await bFixture.c.createCommessa(click);
  assert.equal((await br.collection('commesse').get()).size,1);
  const plant=(user,tenant,index,snowMode=isSnow,parent=job)=>{
   const f=fixture(user,tenant,snowMode,index);
   f.c.managementCommessaId=parent.id;f.c.commesseById.set(parent.id,parent.data());return f;
  };
  for(let i=1;i<=9;i++)await plant(owner,a,i).c.addManagementImpianto();
  const race=await Promise.allSettled([plant(owner,a,10).c.addManagementImpianto(),plant(owner,a,11).c.addManagementImpianto()]);
  assert.equal(race.filter(x=>x.status==='fulfilled').length,1);
  assert.equal(race.find(x=>x.status==='rejected')?.reason?.code,'invalid-quota-plants');
  const ps=await job.ref.collection('impianti').get(),physical=await job.ref.collection('impiantiFisici').get();
  assert.equal(ps.size,10);assert.equal(physical.size,10);
  assert.equal((await ar.get()).data().starterPlan.usage.plants,10);
  for(const p of ps.docs){assert.equal(p.data().done,false);assert.equal(p.data().stato,'DA FARE');assert.equal(p.data().physicalPlantId,p.id);}
  // UI roles are intentionally forged in fixtures: server membership must still deny.
  const invite=await call(owner,'invite',{tenantId:a,recipientEmail:worker.email,role:'worker'});
  await call(worker,'accept',{tenantId:a,inviteToken:invite.token});
  await denied(fixture(worker,a).c.createCommessa(click),'forbidden');
  await denied(plant(other,a,99).c.addManagementImpianto(),'forbidden');
  await denied(plant(worker,a,99).c.addManagementImpianto(),'forbidden');
  await denied(plant(owner,b,99).c.addManagementImpianto(),'forbidden');
  await denied(plant(other,b,99).c.addManagementImpianto(),'invalid-unmanaged-parent');
  assert.equal((await br.get()).data().starterPlan.usage.plants,0);
  assert.throws(()=>normal.scope.db.collection('platformUsers'),{code:'permission-denied'});
  normal.disconnect();assert.throws(()=>normal.scope.db.collection('commesse'),{code:'permission-denied'});
  console.log('Original createCommessa/addManagementImpianto: atomic base-plan limits, dual ordinary/snow, tenant isolation, physical pair and forged UI role denial passed.');
 }finally{for(const s of scopes)s.close();await adapter.close();}
}
main().catch(e=>{console.error(e);process.exitCode=1;});

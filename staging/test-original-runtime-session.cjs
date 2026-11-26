// Actual Firestore compat SDK against the local emulator. No production or live staging data.
const assert=require('node:assert/strict'),fs=require('node:fs');
const {initializeTestEnvironment,assertFails}=require('@firebase/rules-unit-testing');
const {doc,setDoc}=require('firebase/firestore');
(async()=>{
 const endpoint=process.env.FIRESTORE_EMULATOR_HOST;
 assert(endpoint&&/^(127\.0\.0\.1|localhost):\d+$/.test(endpoint),'Local emulator required');
 const {createOriginalRuntimeSession}=await import('./original-runtime-session.mjs');
 // The runtime deliberately only accepts the staging project ID. The testing SDK is
 // explicitly connected to the local emulator, and no live Firebase app is initialized.
 const projectId='varga-cantieri-v2-staging';
 const env=await initializeTestEnvironment({projectId,firestore:{rules:fs.readFileSync(process.argv[2],'utf8')}});
 const runtimes=[];
 try{
 await env.clearFirestore();
 await env.withSecurityRulesDisabled(async ctx=>{
  const db=ctx.firestore();
  for(const [id,uids] of [['demo-a',['u','only-a']],['demo-b',['u','only-b']]]){
   await setDoc(doc(db,'tenants',id),{demo:true,status:'active',name:id});
   for(const uid of uids){
    await setDoc(doc(db,'tenants',id,'members',uid),{role:'admin',status:'active'});
    await setDoc(doc(db,'users',uid,'memberships',id),{tenantId:id,role:'admin',status:'active'});
   }
   await setDoc(doc(db,'tenants',id,'commesse','same-id'),{nome:id});
  }
 });
 const make=uid=>{
  const db=env.authenticatedContext(uid,{email_verified:true,email:uid+'@example.test'}).firestore();
  const user={uid,getIdTokenResult:async()=>({claims:{email_verified:true}})};
  let observer;
  const auth={app:db.app,currentUser:user,onAuthStateChanged(fn){observer=fn;return()=>{};}};
  const runtime=createOriginalRuntimeSession({auth,db});runtimes.push(runtime);
  return {db,auth,runtime,logout(){auth.currentUser=null;observer(null);}};
 };
 const a=make('only-a'),b=make('only-b'),both=make('u');
 const sa=await a.runtime.select('demo-a'),sb=await b.runtime.select('demo-b');
 assert.equal((await sa.firestore.collection('commesse').doc('same-id').get({source:'server'})).data().nome,'demo-a');
 assert.equal((await sb.firestore.collection('commesse').doc('same-id').get({source:'server'})).data().nome,'demo-b');
 await assert.rejects(a.runtime.select('demo-b'));
 await assertFails(a.db.doc('tenants/demo-b').get({source:'server'}));
 await assertFails(a.db.doc('tenants/demo-a/members/u').get({source:'server'}));
 await assertFails(a.db.doc('users/only-b/memberships/demo-b').get({source:'server'}));
 await assertFails(a.db.collection('commesse').get({source:'server'}));
 await assertFails(a.db.doc('tenants/demo-a/members/only-a').update({role:'employee'}));
 const x=await both.runtime.select('demo-a'),y=await both.runtime.select('demo-b');
 assert.throws(()=>x.firestore.collection('commesse'));
 assert.equal((await y.firestore.collection('commesse').doc('same-id').get({source:'server'})).data().nome,'demo-b');
 both.logout();assert.throws(()=>y.firestore.collection('commesse'));
 await env.withSecurityRulesDisabled(async ctx=>setDoc(doc(ctx.firestore(),'tenants','demo-b','members','only-b'),{status:'suspended'},{merge:true}));
 await assert.rejects(b.runtime.select('demo-b'));
 console.log('ORIGINAL RUNTIME SESSION: verified server metadata, isolation, switch, logout, suspension.');
 }finally{for(const runtime of runtimes)runtime.close();await env.cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1;});

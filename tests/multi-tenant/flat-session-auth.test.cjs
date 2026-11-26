'use strict';
// Local Auth+Firestore proof; never expose or import into live functions.
const assert=require('node:assert/strict');
const {randomUUID}=require('node:crypto');
const {readFileSync}=require('node:fs'),{join}=require('node:path');
const {initializeTestEnvironment,assertFails}=require('@firebase/rules-unit-testing');
const {initializeApp,deleteApp}=require('firebase/app');
const {getAuth,connectAuthEmulator,signInWithEmailAndPassword,signInWithCustomToken,setPersistence,inMemoryPersistence}=require('firebase/auth');
const {getFirestore,connectFirestoreEmulator,doc,collection,query,where,orderBy,getDocsFromServer,getDocFromServer,terminate}=require('firebase/firestore');
const {createLocalAdapter,PROJECT}=require('./auth-adapter.emulator.cjs');
const fail=code=>{throw Object.assign(new Error(code),{code});};
async function main(){
 assert.equal(process.env.FIRESTORE_EMULATOR_HOST,'127.0.0.1:8180');
 assert.equal(process.env.FIREBASE_AUTH_EMULATOR_HOST,'127.0.0.1:9199');
 const env=await initializeTestEnvironment({projectId:PROJECT,firestore:{host:'127.0.0.1',port:8180,rules:readFileSync(join(__dirname,'flat-session-experiment.rules'),'utf8')}});
 const adapter=createLocalAdapter(),apps=[];
 const label=randomUUID(),a='auth-a-'+label,b='auth-b-'+label,outsider='auth-out-'+label;
 const job=(tenant)=>'auth-job-'+tenant;
 const sessionPath=(uid,sid)=>'experimentSessions/'+uid+'/devices/'+sid;
 async function client(){
  const app=initializeApp({projectId:PROJECT,apiKey:'demo-key',authDomain:PROJECT+'.firebaseapp.com'},randomUUID());
  const auth=getAuth(app);connectAuthEmulator(auth,'http://127.0.0.1:9199',{disableWarnings:true});
  await setPersistence(auth,inMemoryPersistence);
  const db=getFirestore(app);connectFirestoreEmulator(db,'127.0.0.1',8180);
  apps.push({app,db});return {auth,db};
 }
 async function passwordAccount(verified=true,claims={}){
  const email=randomUUID()+'@example.invalid',password='Test-'+randomUUID();
  const u=await adapter.auth.createUser({email,password,emailVerified:verified});
  await adapter.auth.setCustomUserClaims(u.uid,claims);
  const c=await client();const login=await signInWithEmailAndPassword(c.auth,email,password);
  return {...c,uid:u.uid,user:login.user,token:await login.user.getIdToken()};
 }
 async function issue(token,tenantId){
  // This helper is local-only, not a network handler or a production token issuer.
  let decoded,user;
  try{decoded=await adapter.auth.verifyIdToken(token,true);user=await adapter.auth.getUser(decoded.uid);}
  catch{fail('unauthenticated');}
  if(user.disabled||!user.emailVerified||decoded.email_verified!==true||decoded.email!==user.email||decoded.apple_review===true||user.customClaims?.apple_review===true||decoded.firebase?.sign_in_provider!=='password')fail('unauthenticated');
  const sid=randomUUID(),ref=adapter.db.doc(sessionPath(user.uid,sid));
  // Sign first, return only after atomic membership verification and creation.
  const customToken=await adapter.auth.createCustomToken(user.uid,{vargaExperimentSession:sid});
  await adapter.db.runTransaction(async tx=>{
   const [company,member]=await Promise.all([tx.get(adapter.db.doc('tenants/'+tenantId)),tx.get(adapter.db.doc('tenants/'+tenantId+'/members/'+user.uid))]);
   if(!company.exists||company.data().status!=='active'||!member.exists||member.data().status!=='active'||!['owner','admin'].includes(member.data().role))fail('forbidden');
   tx.create(ref,{tenantId,status:'active',expiresAt:new Date(Date.now()+30*60000)});
  });
  return {sid,customToken};
 }
 const fq=(db,tenant)=>query(collection(db,'commesse'),where('tenantId','==',tenant),where('flatSessionExperiment','==',true),orderBy('createdAt','desc'));
 const read=async(c,tenant)=>{const s=await getDocsFromServer(fq(c.db,tenant));assert.deepEqual(s.docs.map(d=>d.id),[job(tenant)]);};
 async function device(ticket,uid){
  const c=await client();const {user}=await signInWithCustomToken(c.auth,ticket.customToken);
  assert.equal(user.uid,uid);
  const initial=await user.getIdTokenResult();
  assert.equal(initial.claims.vargaExperimentSession,ticket.sid,'Initial ID token lost session claim');
  assert.equal(initial.claims.email_verified,true);
  return {...c,user};
 }
 try{
  const owner=await passwordAccount(),worker=await passwordAccount(),unverified=await passwordAccount(false),review=await passwordAccount(true,{apple_review:true});
  for(const tenant of [a,b,outsider]){
   await adapter.db.doc('tenants/'+tenant).set({status:'active'});
   await adapter.db.doc('commesse/'+job(tenant)).set({tenantId:tenant,flatSessionExperiment:true,nome:tenant,createdAt:new Date()});
   if(tenant!==outsider){
    await adapter.db.doc('tenants/'+tenant+'/members/'+owner.uid).set({status:'active',role:'owner'});
    await adapter.db.doc('tenants/'+tenant+'/members/'+worker.uid).set({status:'active',role:'worker'});
   }
  }
  await assert.rejects(issue(owner.token,outsider),{code:'forbidden'});
  await assert.rejects(issue(worker.token,a),{code:'forbidden'});
  await assert.rejects(issue(unverified.token,a),{code:'unauthenticated'});
  await assert.rejects(issue(review.token,a),{code:'unauthenticated'});
  await assertFails(getDocsFromServer(fq(owner.db,a)));
  const [ticketA,ticketB]=await Promise.all([issue(owner.token,a),issue(owner.token,b)]);
  assert.notEqual(ticketA.sid,ticketB.sid);
  const phone=await device(ticketA,owner.uid),tablet=await device(ticketB,owner.uid);
  await Promise.all([read(phone,a),read(tablet,b)]);
  await assertFails(getDocsFromServer(fq(phone.db,b)));
  await assertFails(getDocsFromServer(fq(tablet.db,a)));
  // A token obtained through a scoped sign-in must not recursively issue sessions.
  await assert.rejects(issue(await phone.user.getIdToken(),b),{code:'unauthenticated'});
  for(const [c,ticket,tenant] of [[phone,ticketA,a],[tablet,ticketB,b]]){
   const refreshed=await c.user.getIdTokenResult(true);
   assert.equal(refreshed.claims.vargaExperimentSession,ticket.sid,'Forced refresh lost session claim');
   await read(c,tenant);
  }
  assert.equal((await adapter.auth.getUser(owner.uid)).customClaims?.vargaExperimentSession,undefined,'Session claim leaked into global user claims');
  await adapter.db.doc(sessionPath(owner.uid,ticketA.sid)).update({status:'revoked'});
  // New client/transport proves server denial independently of existing snapshots.
  const freshPhone=await device(ticketA,owner.uid);
  await assertFails(getDocFromServer(doc(freshPhone.db,'commesse/'+job(a))));
  await assertFails(getDocsFromServer(fq(freshPhone.db,a)));
  await read(tablet,b);
  await adapter.db.doc('tenants/'+b+'/members/'+owner.uid).update({status:'revoked'});
  const freshTablet=await device(ticketB,owner.uid);
  await assertFails(getDocFromServer(doc(freshTablet.db,'commesse/'+job(b))));
  await assertFails(getDocsFromServer(fq(freshTablet.db,b)));
  console.log('Real emulator Auth: password verification, per-session custom token, refresh, scoped access and fresh-request revocation passed.');
 }finally{
  for(const c of apps){try{await terminate(c.db);}finally{await deleteApp(c.app);}}
  await adapter.close();await env.cleanup();
 }
}
main().catch(e=>{console.error('Session Auth proof failed:',e.code||e.name,e.message);process.exitCode=1;});

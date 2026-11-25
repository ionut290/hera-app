// Native Firebase compat SDK + emulator, original source subscription unchanged.
const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const {initializeTestEnvironment,assertFails}=require('@firebase/rules-unit-testing');
const {doc,setDoc,serverTimestamp}=require('firebase/firestore');
const source=fs.readFileSync('../app.js','utf8');
function extract(name){const i=source.indexOf('function '+name+'(');assert(i>=0,name);const e=source.indexOf('\n}\n',i);assert(e>i);return source.slice(i,e+2);}
(async()=>{
 assert(process.env.FIRESTORE_EMULATOR_HOST,'Emulator required');
 const {createOriginalFirestoreCompat}=await import('./original-firestore-compat.mjs');
 const projectId='demo-varga-original-native';
 const env=await initializeTestEnvironment({projectId,firestore:{rules:fs.readFileSync(process.argv[2],'utf8')}});
 const subscriptions=[];
 let count=0;
 try{
 await env.clearFirestore();
 await env.withSecurityRulesDisabled(async ctx=>{const db=ctx.firestore();for(const [t,u] of [['demo-a','a'],['demo-b','b']]){
 await setDoc(doc(db,'tenants',t),{demo:true,status:'active',originalCreateEnabled:true});
 await setDoc(doc(db,'tenants',t,'members',u),{role:'admin',status:'active'});
 await setDoc(doc(db,'users',u,'memberships',t),{tenantId:t,role:'admin',status:'active'});
 }});
 const native=uid=>env.authenticatedContext(uid,{email_verified:true,email:uid+'@example.test'}).firestore();
 const a=native('a'),b=native('b');
 const scoped=(db,t)=>createOriginalFirestoreCompat({db,router:{projectId,context:{tenantId:t},path(n){assert.equal(n,'commesse');return 'tenants/'+t+'/'+n;},close(){}}});
 const ca=scoped(a,'demo-a'),cb=scoped(b,'demo-b');
 function ui(db,uid){
 let signal=()=>{};let failure=null;
 const scope=vm.createContext({db,currentUser:{uid},commesseById:new Map(),commesseLoadState:{},unsubscribeCommesse:null,
 getCommesseCollectionName:()=> 'commesse',renderCommesseHomeList(){},startCommesseLoadTimeout(){},clearCommesseLoadTimeout(){},
 validateFirebaseConfigForCommesse:()=>true,loadCommesseFromLocalCache(){throw Error('Unexpected cache fallback');},
 refreshCommesseDependentUI(){signal();},tryAutoOpenAssignedCommessaAtStartup(){},renderNextActionCard(){},renderTodaySummary(){},saveCommesseLocalCache(){},stopCommessaStatsSubscriptions(){},
 runAfterFirstRender:fn=>fn(),runFirestoreGetWithRetry:q=>q.get({source:'server'}),logFirestoreError:(_,e)=>{failure=e;signal();},getReadableFirestoreError:e=>String(e),console});
 vm.runInContext(['subscribeCommesse','stopCommesseSubscription','normalizeCommessaDocument'].map(extract).join('\n'),scope);
 subscriptions.push(()=>scope.stopCommesseSubscription());
 const untilSize=n=>new Promise((resolve,reject)=>{const timer=setTimeout(()=>{signal=()=>{};reject(Error('Original subscription timeout'));},6000);signal=()=>{if(failure||scope.commesseById.size===n){clearTimeout(timer);signal=()=>{};failure?reject(failure):resolve();}};signal();});
 return {scope,untilSize};
 }
 const ua=ui(ca,'a'),ub=ui(cb,'b');
 await Promise.all([ua.scope.subscribeCommesse(),ub.scope.subscribeCommesse()]);
 assert.equal(ua.scope.commesseLoadState.status,'empty');assert.equal(ub.scope.commesseLoadState.status,'empty');count++;
 const payload=(name,uid)=>({nome:name,codice:'',excelModelVersion:2,priceListVersion:2,workItemsModelVersion:2,percentualeRibassoGenerale:0.01,excelModelActivatedAt:serverTimestamp(),excelModelActivatedBy:uid+'@example.test',nextImpiantoNumber:1,parentCommessaId:null,creatoDa:uid+'@example.test',createdAt:serverTimestamp()});
 const ref=ca.collection('commesse').doc('shared-id');
 await ca.batch().set(ref,payload('A','a')).commit();
 await ua.untilSize(1);assert.equal(ua.scope.commesseById.get('shared-id').nome,'A');count++;
 assert.equal((await cb.collection('commesse').get({source:'server'})).size,0);count++;
 const rb=cb.collection('commesse').doc('shared-id');
 await cb.runTransaction(async tx=>{const snap=await tx.get(rb);assert.equal(snap.exists,false);tx.set(rb,payload('B','b'));});
 await ub.untilSize(1);assert.equal(ub.scope.commesseById.get('shared-id').nome,'B');count++;
 assert.equal((await ca.collection('commesse').doc('shared-id').get({source:'server'})).data().nome,'A');count++;
 const bad=ca.batch();assert.throws(()=>bad.set(rb,payload('bad','a')));await assert.rejects(bad.commit());count++;
 await assert.rejects(ca.runTransaction(async tx=>tx.get(rb)));count++;
 await assertFails(a.doc('tenants/demo-b/commesse/shared-id').get({source:'server'}));count++;
 const forbidden=ca.batch();assert.throws(()=>forbidden.delete(ref));await assert.rejects(forbidden.commit());count++;
 ua.scope.stopCommesseSubscription();assert.equal(ua.scope.unsubscribeCommesse,null);ca.close();assert.throws(()=>ca.collection('commesse'));count++;
 console.log('ORIGINAL NATIVE SDK: '+count+'/10');assert.equal(count,10);
 }finally{for(const stop of subscriptions)stop();await env.cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1});

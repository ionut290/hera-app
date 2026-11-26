// Actual Firestore compat SDK against the local emulator. No production or live staging data.
const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const {initializeTestEnvironment,assertFails}=require('@firebase/rules-unit-testing');
const {doc,setDoc,serverTimestamp}=require('firebase/firestore');
(async()=>{
 const endpoint=process.env.FIRESTORE_EMULATOR_HOST;
 assert(endpoint&&/^(127\.0\.0\.1|localhost):\d+$/.test(endpoint),'Local emulator required');
 const {createOriginalRuntimeSession}=await import('./original-runtime-session.mjs');
 const {createOriginalRuntimeEntry}=await import('./original-runtime-entry.mjs');
 const {buildIntegratedRuntimeCandidate}=await import('./original-runtime-source.mjs');
 const candidate=buildIntegratedRuntimeCandidate(fs.readFileSync('../app.js','utf8'));
 const extract=name=>{const marker='function '+name+'(',i=candidate.indexOf(marker),end=candidate.indexOf('\n}\n',i);assert(i>=0&&end>i);return candidate.slice(candidate.slice(i-6,i)==='async '?i-6:i,end+2);};
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
   await setDoc(doc(db,'tenants',id),{demo:true,status:'active',name:id,originalCreateEnabled:true,originalManualEnabled:true});
   for(const uid of uids){
    await setDoc(doc(db,'tenants',id,'members',uid),{role:'admin',status:'active'});
    await setDoc(doc(db,'users',uid,'memberships',id),{tenantId:id,role:'admin',status:'active'});
   }
   await setDoc(doc(db,'tenants',id,'commesse','same-id'),{nome:id});
  }
 });
 const make=uid=>{
  const db=env.authenticatedContext(uid,{email_verified:true,email:uid+'@example.test'}).firestore();
  const user={uid,email:uid+'@example.test',displayName:'Operatore '+uid,getIdTokenResult:async()=>({claims:{email_verified:true}})};
  const observers=new Set();
  const auth={app:db.app,currentUser:user,onAuthStateChanged(fn){observers.add(fn);return()=>observers.delete(fn);}};
  const runtime=createOriginalRuntimeSession({auth,db});runtimes.push(runtime);
  return {db,auth,runtime,logout(){auth.currentUser=null;for(const observer of [...observers])observer(null);}};
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
 // Connect the integrated candidate's actual original handlers to the entry and
 // the real compat SDK. UI values are synthetic; all database writes use emulator rules.
 const client=make('u'),entry=createOriginalRuntimeEntry({auth:client.auth,db:client.db});runtimes.push(entry);
 const setup=async id=>{
  const prepared=await entry.enter(id),approval=await prepared.approval.verify(client.auth.currentUser);assert(approval.allowed);
  const user=client.auth.currentUser,alerts=[];
  const ui={commessaName:{value:'Commessa integrata '+id},commessaCode:{value:'V2'},commessaType:{value:'main'},commessaParent:{value:''},commessaForm:{reset(){}},commessaTargetSelect:{value:''},manualImpiantoDenominazione:{value:'Parco integrato'},manualImpiantoComune:{value:'Bologna'},manualImpiantoIndirizzo:{value:'Via demo'},manualImpiantoCodice:{value:'A11'},manualImpiantoFeedback:{},manualImpiantoForm:{reset(){}}};
  const scope=vm.createContext({window:{__VARGA_ISOLATED_RUNTIME_CANDIDATE__:prepared},db:prepared.session.firestore,auth:client.auth,currentUser:user,ui,alert:m=>alerts.push(m),driveAccessToken:null,closeNewCommessaDialog(){},updateCommessaParentField(){},getCommesseCollectionName:()=> 'commesse',commesseById:new Map(),selectedCommessaId:'',selectedCommessaName:'',firebase:{firestore:{FieldValue:{serverTimestamp}}}});
  vm.runInContext(['canManageData','createCommessa','addManualImpianto','getTargetCommessaId','getTargetCommessaName','buildImpiantoKey','mergeMultiValue','splitCodes','hasOrdinario','hasStraordinario','classifyTipoManutenzione'].map(extract).join('\n'),scope);
  return {prepared,scope,ui,alerts};
 };
 const first=await setup('demo-a');assert.equal(first.scope.canManageData(),true);
 await first.scope.createCommessa({preventDefault(){}});
 const listA=await first.prepared.session.firestore.collection('commesse').get({source:'server'});
 const newA=listA.docs.find(x=>x.data().nome==='Commessa integrata demo-a');assert(newA);
 first.ui.commessaTargetSelect.value=newA.id;first.scope.commesseById.set(newA.id,newA.data());
 await first.scope.addManualImpianto({preventDefault(){}});
 const plants=await newA.ref.collection('impianti').get({source:'server'});
 assert.equal(plants.size,1);assert.equal(plants.docs[0].data().done,false);assert.equal(plants.docs[0].data().codicePrezzo,'A11');
 const second=await setup('demo-b');assert.equal(first.scope.canManageData(),false);assert.throws(()=>newA.ref.collection('impianti'));
 const listB=await second.prepared.session.firestore.collection('commesse').get({source:'server'});
 assert(!listB.docs.some(x=>x.id===newA.id));assert(!listB.docs.some(x=>x.data().nome==='Commessa integrata demo-a'));
 await second.scope.createCommessa({preventDefault(){}});
 const newB=(await second.prepared.session.firestore.collection('commesse').get({source:'server'})).docs.find(x=>x.data().nome==='Commessa integrata demo-b');assert(newB);
 assert.equal((await newB.ref.collection('impianti').get({source:'server'})).size,0);
 console.log('INTEGRATED ORIGINAL HANDLERS: real emulator create commessa, insert plant, company switch and isolation passed.');
 }finally{for(const runtime of runtimes)runtime.close();await env.cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1;});

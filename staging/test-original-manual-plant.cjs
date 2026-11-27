// Execute the unchanged original manual entry handler against emulator rules.
const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const {initializeTestEnvironment,assertFails}=require('@firebase/rules-unit-testing');
const {doc,collection,setDoc,getDoc,getDocs,addDoc,updateDoc,serverTimestamp}=require('firebase/firestore');
const source=fs.readFileSync('../app.js','utf8');
function extract(name){const token='function '+name+'(';const i=source.indexOf(token);assert(i>=0&&source.indexOf(token,i+token.length)<0,name);const start=source.slice(i-6,i)==='async '?i-6:i;const e=source.indexOf('\n}\n',i);assert(e>i);return source.slice(start,e+2);}
(async()=>{
 assert(process.env.FIRESTORE_EMULATOR_HOST,'Emulator required');
 const {createOriginalFirestoreCompat}=await import('./original-firestore-compat.mjs');
 const projectId='demo-varga-original-manual';
 const env=await initializeTestEnvironment({projectId,firestore:{rules:fs.readFileSync(process.argv[2],'utf8')}});
 let count=0;
 try{
 await env.clearFirestore();
 await env.withSecurityRulesDisabled(async c=>{const db=c.firestore();for(const [t,u] of [['demo-a','a'],['demo-b','b']]){
 await setDoc(doc(db,'tenants',t),{demo:true,status:'active',originalCreateEnabled:true,originalManualEnabled:true});
 for(const [uid,role] of [[u,'admin'],[u+'-worker','employee']]){
 await setDoc(doc(db,'tenants',t,'members',uid),{role,status:'active'});
 await setDoc(doc(db,'users',uid,'memberships',t),{tenantId:t,role,status:'active'});
 }
 await setDoc(doc(db,'tenants',t,'commesse','verde'),{nome:'Verde demo',excelModelVersion:2,nextImpiantoNumber:1});
 }});
 const getDB=uid=>env.authenticatedContext(uid,{email_verified:true,email:uid+'@example.test'}).firestore();
 const make=(uid,t,role='admin')=>{
 const db=getDB(uid);
 const collectionRef=p=>({async get(){return getDocs(collection(db,p));},async add(data){return addDoc(collection(db,p),{...data});},doc(id){assert(!id.includes('/'));const path=p+'/'+id;return {collection(n){assert.equal(n,'impianti');return collectionRef(path+'/'+n);},set:(d,o)=>setDoc(doc(db,path),{...d},{...o}),update:d=>updateDoc(doc(db,path),{...d})};}});
 const router={projectId,context:{tenantId:t},path(n){assert.equal(n,'commesse');return 'tenants/'+t+'/'+n;},close(){}};
 const compat=createOriginalFirestoreCompat({db:{app:{options:{projectId}},collection:collectionRef},router});
 const ui={commessaTargetSelect:{value:'verde'},manualImpiantoDenominazione:{value:'Parco demo'},manualImpiantoComune:{value:'Comune demo'},manualImpiantoIndirizzo:{value:'Via demo'},manualImpiantoCodice:{value:'A11'},manualImpiantoFeedback:{},manualImpiantoForm:{reset(){}}};
 const scope=vm.createContext({ui,auth:{currentUser:{uid}},canManageData:()=>role==='admin',db:compat,getCommesseCollectionName:()=> 'commesse',selectedCommessaId:'verde',selectedCommessaName:'Verde demo',commesseById:new Map([['verde',{nome:'Verde demo'}]]),firebase:{firestore:{FieldValue:{serverTimestamp}}}});
 vm.runInContext(['addManualImpianto','getTargetCommessaId','getTargetCommessaName','buildImpiantoKey','mergeMultiValue','splitCodes','hasOrdinario','hasStraordinario','classifyTipoManutenzione'].map(extract).join('\n'),scope);
 return {db,ui,scope,submit:()=>scope.addManualImpianto({preventDefault(){}})};
 };
 const a=make('a','demo-a'),b=make('b','demo-b'),worker=make('a-worker','demo-a','employee');
 const list=(x,t)=>getDocs(collection(x.db,'tenants',t,'commesse','verde','impianti'));
 const parent=(x,t)=>doc(x.db,'tenants',t,'commesse','verde');
 await a.submit();let snap=await list(a,'demo-a');assert.equal(snap.size,1);count++;
 const first=snap.docs[0];assert.equal(first.data().codicePrezzo,'A11');assert.equal(first.data().done,false);assert.equal(first.data().numeroProgressivo,1);count++;
 assert.equal((await getDoc(parent(a,'demo-a'))).data().nextImpiantoNumber,2);count++;
 a.ui.manualImpiantoCodice.value='B21';await a.submit();snap=await list(a,'demo-a');assert.equal(snap.size,1);assert.equal(snap.docs[0].id,first.id);assert.equal(snap.docs[0].data().codicePrezzo,'A11 | B21');assert.equal(snap.docs[0].data().tipoManutenzione,'Ordinaria + Straordinaria');count++;
 assert.equal((await list(b,'demo-b')).size,0);count++;
 await b.submit();assert.equal((await list(b,'demo-b')).size,1);assert.equal((await list(a,'demo-a')).size,1);count++;
 await worker.submit();assert.match(worker.ui.manualImpiantoFeedback.textContent,/Solo l'admin/);count++;
 await assertFails(setDoc(doc(worker.db,'tenants','demo-a','commesse','verde','impianti','fake'),{...first.data(),createdAt:serverTimestamp()}));count++;
 await assertFails(updateDoc(doc(a.db,first.ref.path),{done:true}));count++;
 await assertFails(updateDoc(doc(a.db,first.ref.path),{denominazione:'changed'}));count++;
 await assertFails(setDoc(doc(a.db,'tenants','demo-b','commesse','verde','impianti','cross'),{...first.data(),createdAt:serverTimestamp()}));count++;
 await assertFails(updateDoc(parent(a,'demo-a'),{nextImpiantoNumber:1000}));count++;
 await env.withSecurityRulesDisabled(async ctx=>setDoc(doc(ctx.firestore(),'tenants','demo-a'),{originalManualEnabled:false},{merge:true}));
 a.ui.manualImpiantoDenominazione.value='Altro parco';await assertFails(a.submit());count++;
 assert.equal((await list(a,'demo-a')).size,1);count++;
 console.log('ORIGINAL MANUAL PLANT: '+count+'/14');assert.equal(count,14);
 }finally{await env.cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1});

// Run original app.js handlers unchanged in a VM with synthetic UI and emulator-only SDK.
// This is not a browser test or a deployable runtime.
const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm');
const {initializeTestEnvironment,assertFails}=require('@firebase/rules-unit-testing');
const {doc,collection,setDoc,getDocs,getDoc,addDoc,serverTimestamp}=require('firebase/firestore');
const source=fs.readFileSync('../app.js','utf8');
function extract(name){const start=source.indexOf('async function '+name+'(');assert(start>=0,name);const end=source.indexOf('\n}\n',start);assert(end>start);return source.slice(start,end+2);}
function extractSync(name){const start=source.indexOf('function '+name+'(');assert(start>=0,name);const end=source.indexOf('\n}\n',start);assert(end>start);return source.slice(start,end+2);}
(async()=>{
 assert(process.env.FIRESTORE_EMULATOR_HOST,'Emulator required');
 const {createOriginalCollectionBridge}=await import('./original-collection-bridge.mjs');
 const env=await initializeTestEnvironment({projectId:'demo-varga-original-create',firestore:{rules:fs.readFileSync(process.argv[2],'utf8')}});
 let count=0;
 try{
 await env.clearFirestore();
 await env.withSecurityRulesDisabled(async ctx=>{const db=ctx.firestore();for(const [id,uid] of [['demo-a','a'],['demo-b','b']]){
 await setDoc(doc(db,'tenants',id),{demo:true,status:'active'});
 for(const [u,role] of [[uid,'admin'],[uid+'-worker','employee']]){
 await setDoc(doc(db,'tenants',id,'members',u),{role,status:'active'});
 await setDoc(doc(db,'users',u,'memberships',id),{tenantId:id,role,status:'active'});
 }}});
 const dbFor=uid=>env.authenticatedContext(uid,{email_verified:true,email:uid+'@example.test'}).firestore();
 const make=(uid,tenant,role='admin')=>{
 const db=dbFor(uid),bridge=createOriginalCollectionBridge({path:n=>{assert.equal(n,'commesse');return 'tenants/'+tenant+'/'+n;},close(){}});
 const ref=path=>({async add(data){return addDoc(collection(db,path),data);},doc(id){assert(!id.includes('/'));return {collection(n){assert.equal(n,'impianti');return ref(path+'/'+id+'/'+n);}};}});
 const user={uid,email:uid+'@example.test'},alerts=[],prompts=['Parco demo','SAP-1'];
 const scope=vm.createContext({auth:{currentUser:user},currentUser:user,ui:{commessaName:{value:'Verde demo'},commessaCode:{value:'V1'},commessaType:{value:'main'},commessaParent:{value:''},commessaForm:{reset(){}}},alert:m=>alerts.push(m),canManageData:()=>role==='admin',db:{collection:p=>{assert.equal(p,'tenants/'+tenant+'/commesse');return ref(p);}},getCommesseCollectionName:()=>bridge.getCommesseCollectionName(false),firebase:{firestore:{FieldValue:{serverTimestamp}}},driveAccessToken:null,closeNewCommessaDialog(){},updateCommessaParentField(){},window:{prompt:()=>prompts.shift()},managementCommessaId:'',commesseById:new Map(),getCommessaCachedImpianti:()=>[],normalizeManagementSearch:v=>String(v).trim().toLowerCase(),getOperatorDisplayName:()=>uid,managementEditingId:'',console});
 vm.runInContext([extract('createCommessa'),extract('addManagementImpianto'),extractSync('validateManagementPlant')].join('\n'),scope);
 return {scope,db,alerts};
 };
 const a=make('a','demo-a'),b=make('b','demo-b'),worker=make('a-worker','demo-a','employee');
 const list=(x,t)=>getDocs(collection(x.db,'tenants',t,'commesse'));
 const click={preventDefault(){}};
 await a.scope.createCommessa(click);
 const listA=await list(a,'demo-a');assert.equal(listA.size,1);count++;
 assert.equal((await list(b,'demo-b')).size,0);count++;
 const parent=listA.docs[0];assert.equal(parent.data().excelModelVersion,2);assert.equal(parent.data().nextImpiantoNumber,1);count++;
 a.scope.managementCommessaId=parent.id;a.scope.commesseById.set(parent.id,parent.data());
 await a.scope.addManagementImpianto();
 const plants=await getDocs(collection(a.db,'tenants','demo-a','commesse',parent.id,'impianti'));assert.equal(plants.size,1);assert.equal(plants.docs[0].data().done,false);assert.equal(plants.docs[0].data().stato,'DA FARE');count++;
 await worker.scope.createCommessa(click);assert.equal(worker.alerts.length,1);assert.equal((await list(a,'demo-a')).size,1);count++;
 worker.scope.managementCommessaId=parent.id;await worker.scope.addManagementImpianto();assert.equal((await getDocs(collection(a.db,'tenants','demo-a','commesse',parent.id,'impianti'))).size,1);count++;
 await assertFails(getDoc(doc(b.db,'tenants','demo-a','commesse',parent.id)));count++;
 await assertFails(setDoc(doc(b.db,'tenants','demo-a','commesse','intrusion'),{...parent.data(),createdAt:serverTimestamp(),excelModelActivatedAt:serverTimestamp(),creatoDa:'b@example.test',excelModelActivatedBy:'b@example.test'}));count++;
 const pdata=plants.docs[0].data();
 await assertFails(setDoc(doc(a.db,'tenants','demo-a','commesse',parent.id,'impianti','fake-done'),{...pdata,done:true,createdAt:serverTimestamp(),updatedAt:serverTimestamp()}));count++;
 await b.scope.createCommessa(click);assert.equal((await list(b,'demo-b')).size,1);assert.equal((await list(a,'demo-a')).size,1);count++;
 console.log('ORIGINAL HANDLERS: '+count+'/10');assert.equal(count,10);
 }finally{await env.cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1});

const {initializeTestEnvironment,assertSucceeds,assertFails}=require('@firebase/rules-unit-testing');
const {doc,setDoc,updateDoc,serverTimestamp,writeBatch}=require('firebase/firestore');
const fs=require('node:fs');
(async()=>{
const env=await initializeTestEnvironment({projectId:'demo-varga-verde-isolation',firestore:{rules:fs.readFileSync(process.argv[2],'utf8')}});
try{
await env.clearFirestore();
await env.withSecurityRulesDisabled(async c=>{
const db=c.firestore();
for(const t of ['demo-a','demo-b','demo-disabled']){
await setDoc(doc(db,'tenants',t),{demo:true,status:'active',coboWriteEnabled:t!=='demo-disabled'});
for(const [uid,role] of [['admin','admin'],['worker','employee']]){
if(t==='demo-b')continue;
await setDoc(doc(db,'tenants',t,'members',uid),{role,status:'active'});
await setDoc(doc(db,'users',uid,'memberships',t),{tenantId:t,role,status:'active'});
}
}
});
const auth=(uid)=>env.authenticatedContext(uid,{email_verified:true,email:uid+'@example.test'}).firestore();
const parent=(db,t)=>doc(db,'tenants',t,'commesse','sfalcio-cobo');
const plant=(db,t,id='cobo-parco-demo')=>doc(db,'tenants',t,'commesse','sfalcio-cobo','impianti',id);
const meta=uid=>({nome:'Sfalcio COBO',codice:'COBO-SFALCIO',speciale:true,tipoSpeciale:'SFALCIO_COBO',verdeBologna:true,parchiGiardini:true,attiva:true,parentCommessaId:null,excelModelVersion:2,priceListVersion:2,workItemsModelVersion:2,percentualeRibassoGenerale:0.01,nextImpiantoNumber:1,creatoDa:uid+'@example.test',createdAt:serverTimestamp()});
const payload=uid=>({id:'cobo-parco-demo',commessaId:'sfalcio-cobo',sfalcioCobo:true,updatedByUid:uid,updatedAt:serverTimestamp(),stato:'da_fare',done:false,doneAt:null,doneBy:'',createdAt:serverTimestamp()});
const a=auth('admin'),w=auth('worker');
await assertFails(setDoc(parent(w,'demo-a'),meta('worker')));
await assertFails(setDoc(parent(a,'demo-disabled'),meta('admin')));
await assertFails(setDoc(parent(a,'demo-b'),meta('admin')));
const batch=writeBatch(a);batch.set(parent(a,'demo-a'),meta('admin'));batch.set(plant(a,'demo-a'),payload('admin'));await assertSucceeds(batch.commit());
await assertSucceeds(updateDoc(plant(w,'demo-a'),{note:'demo',updatedByUid:'worker',updatedAt:serverTimestamp()}));
await assertFails(updateDoc(plant(w,'demo-a'),{done:true,updatedByUid:'worker',updatedAt:serverTimestamp()}));
await assertFails(updateDoc(plant(w,'demo-a'),{note:'demo',updatedByUid:'admin',updatedAt:serverTimestamp()}));
await assertFails(setDoc(plant(a,'demo-b'),payload('admin')));
console.log('COBO: 8 checks passed');
}finally{await env.cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1;});

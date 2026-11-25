const fs=require('node:fs');
const {initializeTestEnvironment,assertSucceeds,assertFails}=require('@firebase/rules-unit-testing');
const {doc,setDoc,getDoc,deleteDoc,serverTimestamp}=require('firebase/firestore');
(async()=>{
const env=await initializeTestEnvironment({projectId:'demo-varga-verde-isolation',firestore:{rules:fs.readFileSync(process.argv[2],'utf8')}});
let count=0;
const check=async(p,allowed)=>{await (allowed?assertSucceeds(p):assertFails(p));count++};
try{
await env.clearFirestore();
await env.withSecurityRulesDisabled(async c=>{
const db=c.firestore();
for(const [id,mode] of [['demo-a','configurable-demo'],['demo-b','configurable-demo'],['demo-protected','immutable']]){
await setDoc(doc(db,'tenants',id),{demo:true,status:'active',fattoMode:mode});
for(const [uid,role] of [['admin','admin'],['employee','employee']]){
if(id==='demo-b')continue;
await setDoc(doc(db,'tenants',id,'members',uid),{status:'active',role});
await setDoc(doc(db,'users',uid,'memberships',id),{status:'active',role,tenantId:id});
}
}
});
const db=uid=>env.authenticatedContext(uid,{email_verified:true}).firestore();
const ref=(uid,id='demo-a')=>doc(db(uid),'tenants',id,'fattoSettings','workflow');
const payload=(version=1,actions=['saveCompletion'])=>({version,actions,updatedAt:serverTimestamp(),updatedBy:'admin'});
await check(setDoc(ref('admin'),payload()),true);
await check(getDoc(ref('admin')),true);
await check(setDoc(ref('admin'),payload(2,['saveCompletion','moveToCompleted'])),true);
await check(setDoc(ref('admin'),payload(2)),false);
await check(setDoc(ref('admin'),payload(4)),false);
for(const actions of [[],['openWhatsApp','saveCompletion'],['saveCompletion','openWhatsApp'],['saveCompletion','exportSheet'],['saveCompletion','saveCompletion'],['saveCompletion','eval']])await check(setDoc(ref('admin'),payload(3,actions)),false);
await check(setDoc(ref('admin','demo-b'),payload()),false);
await check(setDoc(ref('admin','demo-protected'),payload()),false);
await check(setDoc(ref('employee'),{...payload(3),updatedBy:'employee'}),false);
await check(getDoc(ref('employee')),false);
await check(deleteDoc(ref('admin')),false);
await check(setDoc(ref('admin'),{...payload(3),updatedBy:'other'}),false);
await check(setDoc(ref('admin'),{...payload(3),extra:true}),false);
console.log('FATTO RESULT: '+count+'/18');
if(count!==18)throw Error('Unexpected count');
}finally{await env.cleanup()}
})().catch(e=>{console.error(e);process.exitCode=1});

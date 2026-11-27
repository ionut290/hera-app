// Test independent price lists, same code/IDs in different companies and contracts.
const assert=require('node:assert/strict'),fs=require('node:fs');
const {initializeTestEnvironment,assertSucceeds,assertFails}=require('@firebase/rules-unit-testing');
const {doc,collection,getDoc,getDocs,setDoc,updateDoc,deleteDoc,serverTimestamp}=require('firebase/firestore');
(async()=>{
 assert(process.env.FIRESTORE_EMULATOR_HOST,'Emulator required');
 const env=await initializeTestEnvironment({projectId:'demo-varga-price-lists',firestore:{rules:fs.readFileSync(process.argv[2],'utf8')}});
 let count=0;
 try{
 await env.clearFirestore();
 await env.withSecurityRulesDisabled(async ctx=>{
 const db=ctx.firestore();
 for(const [t,u,enabled] of [['demo-a','a',true],['demo-b','b',true],['demo-off','a',false]]){
 await setDoc(doc(db,'tenants',t),{demo:true,status:'active',priceListEnabled:enabled,priceListWriteEnabled:enabled});
 for(const [uid,role] of [[u,'admin'],[u+'-worker','employee']]){
 await setDoc(doc(db,'tenants',t,'members',uid),{role,status:'active'});
 await setDoc(doc(db,'users',uid,'memberships',t),{tenantId:t,role,status:'active'});
 }
 for(const c of ['verde','neve'])await setDoc(doc(db,'tenants',t,'commesse',c),{nome:c,percentualeRibassoGenerale:0});
 }});
 const db=uid=>env.authenticatedContext(uid,{email_verified:true,email:uid+'@example.test'}).firestore();
 const ref=(uid,t,c,id='A11')=>doc(db(uid),'tenants',t,'commesse',c,'prezziario',id);
 const parent=(uid,t,c)=>doc(db(uid),'tenants',t,'commesse',c);
 const actor=uid=>({updatedAt:serverTimestamp(),updatedBy:uid,updatedByName:'Operatore demo'});
 const data=(uid,c,p)=>({commessaId:c,numeroProgressivo:1,codiceVoce:'A11',descrizione:'Sfalcio',unitaMisura:'mq',prezzoBase:p,percentualeRibasso:null,prezzoRibassato:p,note:'',createdAt:serverTimestamp(),createdBy:uid,...actor(uid)});
 const run=async(p,ok)=>{await(ok?assertSucceeds(p):assertFails(p));count++;};
 for(const [u,t,c] of [['a','demo-a','verde'],['a','demo-a','neve'],['b','demo-b','verde']]){
 assert.equal((await getDocs(collection(db(u),'tenants',t,'commesse',c,'prezziario'))).size,0);count++;
 }
 await run(setDoc(ref('a','demo-a','verde'),data('a','verde',12)),true);
 await run(setDoc(ref('a','demo-a','neve'),data('a','neve',19)),true);
 await run(setDoc(ref('b','demo-b','verde'),data('b','verde',32)),true);
 for(const [u,t,c,p] of [['a','demo-a','verde',12],['a','demo-a','neve',19],['b','demo-b','verde',32]]){assert.equal((await getDoc(ref(u,t,c))).data().prezzoBase,p);count++;}
 await run(updateDoc(ref('a','demo-a','verde'),{prezzoBase:15,prezzoRibassato:15,...actor('a')}),true);
 assert.equal((await getDoc(ref('b','demo-b','verde'))).data().prezzoBase,32);count++;
 assert.equal((await getDoc(ref('a','demo-a','neve'))).data().prezzoBase,19);count++;
 await run(updateDoc(parent('a','demo-a','verde'),{percentualeRibassoGenerale:0.15,...actor('a')}),true);
 assert.equal((await getDoc(parent('b','demo-b','verde'))).data().percentualeRibassoGenerale,0);count++;
 assert.equal((await getDoc(parent('a','demo-a','neve'))).data().percentualeRibassoGenerale,0);count++;
 await run(getDoc(ref('a-worker','demo-a','verde')),true);
 await run(getDoc(ref('b','demo-a','verde')),false);
 await run(setDoc(ref('a','demo-b','verde','new'),data('a','verde',1)),false);
 await run(setDoc(ref('a-worker','demo-a','verde','new'),data('a-worker','verde',1)),false);
 await run(setDoc(ref('a','demo-a','missing'),data('a','missing',1)),false);
 await run(setDoc(ref('a','demo-a','verde','wrong-parent'),data('a','neve',1)),false);
 await run(setDoc(ref('a','demo-a','verde','negative'),data('a','verde',-1)),false);
 await run(updateDoc(ref('a','demo-a','verde'),{percentualeRibasso:1.1,...actor('a')}),false);
 await run(updateDoc(ref('a','demo-a','verde'),{done:true,...actor('a')}),false);
 await run(updateDoc(ref('a','demo-a','verde'),{createdBy:'b',...actor('a')}),false);
 await run(deleteDoc(ref('a','demo-a','verde')),false);
 await run(setDoc(ref('a','demo-off','verde'),data('a','verde',1)),false);
 await run(updateDoc(parent('a','demo-a','verde'),{percentualeRibassoGenerale:1.5,...actor('a')}),false);
 await run(getDocs(collection(db('a'),'commesse','verde','prezziario')),false);
 assert.equal(count,29);console.log('PRICE LIST ISOLATION: '+count+'/29');
 }finally{await env.cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1});

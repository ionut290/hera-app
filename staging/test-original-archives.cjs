// Native Firestore emulator tests. Synthetic data only, no production.
const assert=require('node:assert/strict'),fs=require('node:fs');
const {initializeTestEnvironment,assertSucceeds,assertFails}=require('@firebase/rules-unit-testing');
const {doc,collection,setDoc,getDoc,getDocs,updateDoc,deleteDoc}=require('firebase/firestore');
(async()=>{
 assert(process.env.FIRESTORE_EMULATOR_HOST,'Emulator required');
 const env=await initializeTestEnvironment({projectId:'demo-varga-original-archives',firestore:{rules:fs.readFileSync(process.argv[2],'utf8')}});
 const paths=['personale','mezzi','neve_personale','neve_mezzi','squadreCommesse','squadreStorico','neve_squadre','neve_squadre_storico','oreReports','oreApprovalRequests','neve_ore','neve_ore_richieste','commessaResources','oreLocks','neve_commesse','neve_commesse/c1/impianti','globalCommesse','globalCommesse/c1/impianti'];
 let count=0;
 try{
 await env.clearFirestore();
 await env.withSecurityRulesDisabled(async ctx=>{
 const db=ctx.firestore();
 for(const [tenant,enabled,status,demo] of [['demo-a',true,'active',true],['demo-b',true,'active',true],['demo-off',false,'active',true],['demo-suspended',true,'suspended',true],['real-a',true,'active',false]]){
 await setDoc(doc(db,'tenants',tenant),{demo,status,originalArchivesEnabled:enabled});
 const users=tenant==='demo-b'?[['b','admin']]:[['a','admin'],['worker','employee'],['supervisor','supervisor']];
 for(const [uid,role] of users){
 await setDoc(doc(db,'tenants',tenant,'members',uid),{status:'active',role});
 await setDoc(doc(db,'users',uid,'memberships',tenant),{tenantId:tenant,status:'active',role});
 }
 for(const p of paths)await setDoc(doc(db,'tenants',tenant,...p.split('/'),'sample'),{name:tenant});
 }
 for(const [id,data] of Object.entries({public:{type:'ferie'},private:{type:'SCADENZA_DOCUMENTO',ownerUserId:'a'},shared:{type:'SCADENZA_DOCUMENTO',sharedToAll:true},authorized:{type:'SCADENZA_DOCUMENTO',authorizedUserIds:['worker']}}))await setDoc(doc(db,'tenants','demo-a','calendarEvents',id),data);
 await setDoc(doc(db,'tenants','demo-a','programmazioni','assigned'),{operatoriCoinvolti:['worker@example.test']});
 await setDoc(doc(db,'tenants','demo-a','programmazioni','unassigned'),{operatoriCoinvolti:['a@example.test']});
 });
 const db=(uid,claims={email_verified:true,email:uid+'@example.test'})=>env.authenticatedContext(uid,claims).firestore();
 const run=async(p,ok)=>{await (ok?assertSucceeds(p):assertFails(p));count++;};
 for(const p of paths){
 const ref=(d,t,id='sample')=>doc(d,'tenants',t,...p.split('/'),id);
 for(const uid of ['a','worker','supervisor']){
 await run(getDoc(ref(db(uid),'demo-a')),true);
 await run(getDocs(collection(db(uid),'tenants','demo-a',...p.split('/'))),true);
 }
 for(const [d,t] of [[db('b'),'demo-a'],[db('a'),'demo-b'],[db('a'),'demo-off'],[db('a'),'demo-suspended'],[db('a'),'real-a'],[db('a',{email_verified:false}),'demo-a'],[env.unauthenticatedContext().firestore(),'demo-a']])await run(getDoc(ref(d,t)),false);
 await run(setDoc(ref(db('a'),'demo-a','new'),{name:'new'}),false);
 await run(updateDoc(ref(db('a'),'demo-a'),{name:'changed'}),false);
 await run(deleteDoc(ref(db('a'),'demo-a')),false);
 await run(getDocs(collection(db('a'),...p.split('/'))),false);
 }
 const cfg=(uid,t,id)=>doc(db(uid),'tenants',t,'appConfig',id);
 await env.withSecurityRulesDisabled(async ctx=>{for(const t of ['demo-a','demo-b'])for(const id of ['activeCommesse','squadreView','neveSquadreView','driveAdminSecret'])await setDoc(doc(ctx.firestore(),'tenants',t,'appConfig',id),{name:t});});
 for(const id of ['activeCommesse','squadreView','neveSquadreView']){
 await run(getDoc(cfg('worker','demo-a',id)),true);
 await run(getDoc(cfg('b','demo-a',id)),false);
 await run(updateDoc(cfg('a','demo-a',id),{name:'changed'}),false);
 }
 await run(getDoc(cfg('a','demo-a','driveAdminSecret')),false);
 await run(getDocs(collection(db('a'),'tenants','demo-a','appConfig')),false);
 const event=(uid,id)=>doc(db(uid),'tenants','demo-a','calendarEvents',id);
 await run(getDoc(event('worker','public')),true);
 await run(getDoc(event('worker','private')),false);
 await run(getDoc(event('a','private')),true);
 await run(getDoc(event('worker','shared')),true);
 await run(getDoc(event('worker','authorized')),true);
 await run(getDoc(event('b','public')),false);
 const program=(uid,id)=>doc(db(uid),'tenants','demo-a','programmazioni',id);
 await run(getDoc(program('worker','assigned')),true);
 await run(getDoc(program('worker','unassigned')),false);
 await run(getDoc(program('a','unassigned')),true);
 await run(getDoc(program('b','assigned')),false);
 console.log('ORIGINAL ARCHIVES READ ONLY: '+count+'/327');
 assert.equal(count,327);
 }finally{await env.cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1;});

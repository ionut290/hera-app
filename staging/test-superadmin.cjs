const assert=require('node:assert/strict'),fs=require('node:fs');
const {initializeTestEnvironment,assertSucceeds,assertFails}=require('@firebase/rules-unit-testing');
const {doc,collection,getDoc,getDocs,setDoc,updateDoc,deleteDoc,writeBatch,serverTimestamp}=require('firebase/firestore');
(async()=>{
 assert.match(process.env.FIRESTORE_EMULATOR_HOST||'',/^(127\.0\.0\.1|localhost):\d+$/);
 const env=await initializeTestEnvironment({projectId:'demo-varga-superadmin',firestore:{rules:fs.readFileSync(process.argv[2],'utf8')}});
 let n=0;
 const run=async(p,ok)=>{await(ok?assertSucceeds(p):assertFails(p));n++;};
 try{
  await env.clearFirestore();
  await env.withSecurityRulesDisabled(async ctx=>{
   const db=ctx.firestore();
   for(const tid of ['demo-alpha','demo-beta','real-one']){
    await setDoc(doc(db,'tenants',tid),{name:tid,status:'active',demo:tid.startsWith('demo-'),originalArchivesEnabled:true});
    await setDoc(doc(db,'tenants',tid,'commesse','same-id'),{nome:tid});
    for(const uid of ['worker','root']){
     await setDoc(doc(db,'tenants',tid,'members',uid),{role:'employee',status:'active',demo:true,historicalId:'preserved-'+uid});
     await setDoc(doc(db,'users',uid,'memberships',tid),{tenantId:tid,role:'employee',status:'active',historicalId:'preserved-'+uid});
    }
   }
   await setDoc(doc(db,'tenants','demo-alpha','members','non-demo'),{role:'employee',status:'active',demo:false});
   await setDoc(doc(db,'users','non-demo','memberships','demo-alpha'),{tenantId:'demo-alpha',role:'employee',status:'active'});
  });
  const client=(uid,claims={})=>env.authenticatedContext(uid,{email_verified:true,...claims}).firestore();
  const root=client('root',{superAdmin:true}),ordinary=client('worker'),unverified=client('root',{superAdmin:true,email_verified:false}),review=client('root',{superAdmin:true,apple_review:true}),outsider=client('outsider',{superAdmin:true});
  const tenant=(db,tid='demo-alpha')=>doc(db,'tenants',tid);
  const member=(db,uid='worker',tid='demo-alpha')=>doc(db,'tenants',tid,'members',uid);
  const link=(db,uid='worker',tid='demo-alpha')=>doc(db,'users',uid,'memberships',tid);
  const change=(db,uid,tid,patch)=>{const b=writeBatch(db);b.update(member(db,uid,tid),patch);b.update(link(db,uid,tid),patch);return b.commit();};
  const data=()=>({name:'Demo new',status:'active',demo:true,createdAt:serverTimestamp()});
  await run(getDocs(collection(root,'tenants')),true);
  await run(getDocs(collection(root,'tenants','demo-alpha','members')),true);
  await run(getDocs(collection(ordinary,'tenants')),false);
  await run(getDocs(collection(unverified,'tenants')),false);
  await run(getDocs(collection(review,'tenants')),false);
  await run(setDoc(tenant(root,'demo-new'),data()),true);
  await run(setDoc(tenant(ordinary,'demo-denied'),data()),false);
  await run(setDoc(tenant(root,'demo-bad'),{...data(),demo:false}),false);
  await run(setDoc(tenant(root,'demo-bad'),{...data(),originalCreateEnabled:true}),false);
  await run(updateDoc(tenant(root),{name:'Renamed'}),true);
  await run(updateDoc(tenant(root),{status:'suspended'}),true);
  await run(getDoc(doc(ordinary,'tenants','demo-alpha','commesse','same-id')),false);
  await run(updateDoc(tenant(root),{status:'active'}),true);
  await run(updateDoc(tenant(root,'real-one'),{name:'Forbidden'}),false);
  await run(updateDoc(tenant(root),{demo:false}),false);
  await run(updateDoc(tenant(root),{originalCreateEnabled:true}),false);
  await run(deleteDoc(tenant(root)),false);
  await run(updateDoc(member(root),{role:'admin'}),false);
  await run(updateDoc(link(root),{role:'admin'}),false);
  await run(change(root,'worker','demo-alpha',{role:'supervisor'}),true);
  await run(change(root,'worker','demo-alpha',{status:'suspended'}),true);
  await run(getDoc(doc(ordinary,'tenants','demo-alpha','commesse','same-id')),false);
  await run(change(root,'worker','demo-alpha',{status:'active'}),true);
  await run(change(ordinary,'worker','demo-alpha',{role:'admin'}),false);
  await run(change(root,'root','demo-alpha',{role:'admin'}),false);
  await run(change(root,'worker','real-one',{role:'admin'}),false);
  await run(change(root,'non-demo','demo-alpha',{role:'admin'}),false);
  await run(change(root,'worker','demo-alpha',{historicalId:'modified'}),false);
  await run(change(root,'worker','demo-alpha',{role:'superAdmin'}),false);
  await run(deleteDoc(member(root)),false);
  await run(deleteDoc(link(root)),false);
  await run(getDoc(doc(outsider,'tenants','demo-alpha','commesse','same-id')),false);
  await run(setDoc(member(root,'new-user'),{role:'admin',status:'active',demo:true}),false);
  assert.equal((await getDoc(member(root))).data().historicalId,'preserved-worker');
  assert.equal((await getDoc(link(root))).data().historicalId,'preserved-worker');
  assert.equal((await getDoc(member(root,'worker','demo-beta'))).data().role,'employee');
  assert.equal(n,33);
  console.log('SUPER ADMIN METADATA: '+n+'/33 permissions, mirrored roles/status and preserved historical identifiers.');
 }finally{await env.cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1;});

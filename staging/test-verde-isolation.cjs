// Emulator-only tests. Run from the existing staging setup with its dependencies.
// Pass the candidate full rules file as argv[2].
const { initializeTestEnvironment, assertSucceeds, assertFails } = require('@firebase/rules-unit-testing');
const { doc, setDoc, getDoc, getDocs, collection, updateDoc, deleteDoc } = require('firebase/firestore');
const fs = require('node:fs');
const assert = require('node:assert/strict');
(async () => {
  assert(process.env.FIRESTORE_EMULATOR_HOST, 'Emulator required');
  const env = await initializeTestEnvironment({projectId:'demo-varga-verde-isolation',firestore:{rules:fs.readFileSync(process.argv[2],'utf8')}});
  const names = ['verdeLevatoRecords','verdeLevatoCommesse','verdeLevatoConfig'];
  let passed=0;
  try {
    await env.clearFirestore();
    await env.withSecurityRulesDisabled(async ctx => {
      const db=ctx.firestore();
      for(const [id,demo,status] of [['demo-a',true,'active'],['demo-b',true,'active'],['demo-off',true,'suspended'],['real-a',false,'active']]){
        await setDoc(doc(db,'tenants',id),{demo,status});
        for(const role of ['admin','supervisor','employee']) {
          await setDoc(doc(db,'tenants',id,'members',role),{role,status:'active'});
          await setDoc(doc(db,'users',role,'memberships',id),{tenantId:id,role,status:'active'});
        }
        await setDoc(doc(db,'tenants',id,'members','blocked'),{role:'employee',status:'suspended'});
        for(const name of names) await setDoc(doc(db,'tenants',id,name,'sample'),{demo:true});
      }
      await setDoc(doc(db,'tenants','demo-b','members','outsider'),{role:'employee',status:'active'});
      await setDoc(doc(db,'users','outsider','memberships','demo-b'),{tenantId:'demo-b',role:'employee',status:'active'});
      for(const [uid,role,membership] of [
        ['missing','employee',null],
        ['mismatch','employee',{tenantId:'demo-a',role:'admin',status:'active'}],
        ['invalid','unknown',{tenantId:'demo-a',role:'unknown',status:'active'}],
        ['inactive','employee',{tenantId:'demo-a',role:'employee',status:'suspended'}],
        ['wrongtenant','employee',{tenantId:'demo-b',role:'employee',status:'active'}]
      ]){
        await setDoc(doc(db,'tenants','demo-a','members',uid),{role,status:'active'});
        if(membership) await setDoc(doc(db,'users',uid,'memberships','demo-a'),membership);
      }
    });
    const db=(uid,claims={email_verified:true})=>env.authenticatedContext(uid,claims).firestore();
    const run=async(label,promise,allowed)=>{await (allowed?assertSucceeds(promise):assertFails(promise));passed++;console.log('PASS '+label)};
    for(const name of names){
      const ref=(d,id='demo-a')=>doc(d,'tenants',id,name,'sample');
      for(const role of ['admin','supervisor','employee']){
        await run(name+' '+role+' read',getDoc(ref(db(role))),true);
        await run(name+' '+role+' list',getDocs(collection(db(role),'tenants','demo-a',name)),true);
        await run(name+' '+role+' update',updateDoc(ref(db(role)),{demo:false}),false);
        await run(name+' '+role+' create',setDoc(doc(db(role),'tenants','demo-a',name,'new'),{demo:true}),false);
        await run(name+' '+role+' delete',deleteDoc(ref(db(role))),false);
      }
      for(const [label,d,id] of [['cross tenant',db('outsider'),'demo-a'],['suspended member',db('blocked'),'demo-a'],['suspended tenant',db('admin'),'demo-off'],['real tenant',db('admin'),'real-a'],['unverified',db('admin',{email_verified:false}),'demo-a'],['apple review',db('admin',{email_verified:true,apple_review:true}),'demo-a'],['anonymous',env.unauthenticatedContext().firestore(),'demo-a']]){
        await run(name+' '+label+' get',getDoc(ref(d,id)),false);
        await run(name+' '+label+' list',getDocs(collection(d,'tenants',id,name)),false);
      }
      for(const uid of ['missing','mismatch','invalid','inactive','wrongtenant']){
        await run(name+' '+uid+' get',getDoc(ref(db(uid))),false);
        await run(name+' '+uid+' list',getDocs(collection(db(uid),'tenants','demo-a',name)),false);
      }
      await run(name+' root denied',getDocs(collection(db('admin'),name)),false);
    }
    console.log('RESULT: '+passed+'/120');
  } finally {await env.cleanup();}
})().catch(e=>{console.error(e);process.exitCode=1});

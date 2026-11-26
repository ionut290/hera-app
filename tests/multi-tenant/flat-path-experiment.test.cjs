'use strict';
// Proof of Firestore path isolation only. Does not execute, wrap or replace FATTO.
const assert=require('node:assert/strict');
const {readFileSync}=require('node:fs'),{join}=require('node:path');
const {initializeTestEnvironment,assertFails,assertSucceeds}=require('@firebase/rules-unit-testing');
const {doc,collection,setDoc,getDocFromServer,getDocsFromServer,query,where,orderBy,updateDoc,writeBatch}=require('firebase/firestore');
async function main(){
  assert.equal(process.env.FIRESTORE_EMULATOR_HOST,'127.0.0.1:8180');
  assert.equal(process.env.FIREBASE_AUTH_EMULATOR_HOST,'127.0.0.1:9199');
  const env=await initializeTestEnvironment({projectId:'demo-varga-multi-tenant',firestore:{host:'127.0.0.1',port:8180,rules:readFileSync(join(__dirname,'flat-path-experiment.rules'),'utf8')}});
  const a='flat-a',b='flat-b',uid='flat-shared-owner',worker='flat-worker',areas=['commesse','neve_commesse'],children=['impianti','impiantiFisici','lavorazioni'];
  const job=(area,tenant)=>'flat-'+area+'-'+tenant;
  const seed=(p,d)=>env.withSecurityRulesDisabled(ctx=>setDoc(doc(ctx.firestore(),p),d));
  try{
    await env.withSecurityRulesDisabled(async ctx=>{
      const db=ctx.firestore(),put=(p,d)=>setDoc(doc(db,p),d);
      for(const tenant of [a,b]){
        await put('tenants/'+tenant,{status:'active'});
        await put('tenants/'+tenant+'/members/'+uid,{status:'active',role:'owner'});
        await put('tenants/'+tenant+'/members/'+worker,{status:'active',role:'worker'});
        for(const area of areas){
          const base=area+'/'+job(area,tenant);
          await put(base,{tenantId:tenant,flatPathExperiment:true,nome:tenant+' '+area,createdAt:new Date()});
          for(const child of children)await put(base+'/'+child+'/same-id',{flatPathExperiment:true,experimentMarker:tenant,stato:'DA FARE'});
        }
      }
      await put('experimentActiveCompany/'+uid,{tenantId:a});
      await put('experimentActiveCompany/'+worker,{tenantId:a});
    });
    const db=env.authenticatedContext(uid,{email_verified:true}).firestore();
    const fw=env.authenticatedContext(worker,{email_verified:true}).firestore();
    const fr=env.authenticatedContext(uid,{email_verified:true,apple_review:true}).firestore();
    const fq=(area,tenant)=>query(collection(db,area),where('tenantId','==',tenant),where('flatPathExperiment','==',true),orderBy('createdAt','desc'));
    for(const area of areas){
      const own=area+'/'+job(area,a),other=area+'/'+job(area,b);
      await assertFails(getDocsFromServer(query(collection(db,area),orderBy('createdAt','desc'))));
      const q=await assertSucceeds(getDocsFromServer(fq(area,a)));assert.deepEqual(q.docs.map(d=>d.id),[job(area,a)]);
      await assertFails(getDocsFromServer(fq(area,b)));
      await assertSucceeds(getDocFromServer(doc(db,own)));
      await assertFails(getDocFromServer(doc(db,other)));
      for(const child of children){
        await assertSucceeds(getDocFromServer(doc(db,own+'/'+child+'/same-id')));
        await assertFails(getDocFromServer(doc(db,other+'/'+child+'/same-id')));
        await assertFails(getDocFromServer(doc(fw,own+'/'+child+'/same-id')));
        await assertFails(getDocFromServer(doc(fr,own+'/'+child+'/same-id')));
      }
      const batch=writeBatch(db);for(const child of children)batch.update(doc(db,own+'/'+child+'/same-id'),{experimentMarker:'local proof'});await assertSucceeds(batch.commit());
      const mixed=writeBatch(db);mixed.update(doc(db,own+'/impianti/same-id'),{experimentMarker:'must not persist'});mixed.update(doc(db,other+'/impianti/same-id'),{experimentMarker:'forbidden'});await assertFails(mixed.commit());
      assert.equal((await getDocFromServer(doc(db,own+'/impianti/same-id'))).data().experimentMarker,'local proof');
      await assertFails(updateDoc(doc(db,own+'/impianti/same-id'),{stato:'FATTO'}));
      await assertFails(setDoc(doc(db,area+'/flat-forbidden-create'),{tenantId:a,flatPathExperiment:true}));
      await assertFails(updateDoc(doc(db,own),{tenantId:b}));
    }
    await assertFails(setDoc(doc(db,'experimentActiveCompany/'+uid),{tenantId:b}));
    // Switch only through a trusted test fixture. Old tenant paths fail immediately.
    await seed('experimentActiveCompany/'+uid,{tenantId:b});
    for(const area of areas){await assertFails(getDocsFromServer(fq(area,a)));assert.equal((await assertSucceeds(getDocsFromServer(fq(area,b)))).size,1);await assertFails(getDocFromServer(doc(db,area+'/'+job(area,a)+'/impianti/same-id')));}
    await seed('tenants/'+b+'/members/'+uid,{status:'revoked',role:'owner'});
    await assertFails(getDocsFromServer(fq('commesse',b)));
    console.log('Flat legacy path experiment: scoped queries, direct child access, atomic mixed-batch denial, selection and revocation passed; FATTO writes denied.');
  }finally{await env.cleanup();}
}
main().catch(e=>{console.error(e);process.exitCode=1;});

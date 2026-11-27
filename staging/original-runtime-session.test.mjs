import test from 'node:test';import assert from 'node:assert/strict';
import {createOriginalRuntimeSession} from './original-runtime-session.mjs';
function fixture(){
 const app={options:{projectId:'varga-cantieri-v2-staging'}};
 const user={uid:'u',getIdTokenResult:async()=>({claims:{email_verified:true}})};
 const docs=new Map();for(const id of ['demo-a','demo-b']){docs.set('tenants/'+id,{demo:true,status:'active'});docs.set('tenants/'+id+'/members/u',{role:'admin',status:'active'});docs.set('users/u/memberships/'+id,{role:'admin',status:'active',tenantId:id});}
 let observer,unsubscribed=false;const calls=[];
 const auth={app,currentUser:user,onAuthStateChanged(fn){observer=fn;return()=>{unsubscribed=true;};}};
 const db={app,doc(path){return {path,get:async opt=>{assert.equal(opt.source,'server');return {exists:docs.has(path),data:()=>docs.get(path)};}};},collection(path){calls.push(path);return {path};}};
 return {auth,db,docs,calls,emit(u){auth.currentUser=u;observer(u);},unsubscribed:()=>unsubscribed};
}
test('original compat runtime verifies membership then scopes original collection names',async()=>{const f=fixture(),runtime=createOriginalRuntimeSession(f),a=await runtime.select('demo-a');assert.equal(a.firestore.collection('commesse').path,'tenants/demo-a/commesse');assert.equal(a.firestore.collection('neve_commesse').path,'tenants/demo-a/neve_commesse');assert.throws(()=>a.firestore.collection('tenants/demo-b/commesse'));assert.throws(()=>a.firestore.collection('driveAdminSecret'));const b=await runtime.select('demo-b');assert.throws(()=>a.firestore.collection('commesse'));assert.equal(b.firestore.collection('commesse').path,'tenants/demo-b/commesse');runtime.close();assert(f.unsubscribed());assert.throws(()=>b.firestore.collection('commesse'));});
test('rejects production and mismatched SDK apps',()=>{const f=fixture();f.db.app={options:{projectId:'hera-app-6cd2b'}};assert.throws(()=>createOriginalRuntimeSession(f));});
test('denies missing membership and invalidates on logout',async()=>{const f=fixture(),runtime=createOriginalRuntimeSession(f);f.docs.delete('users/u/memberships/demo-b');await assert.rejects(runtime.select('demo-b'));const a=await runtime.select('demo-a');f.emit(null);assert.throws(()=>a.firestore.collection('commesse'));await assert.rejects(runtime.select('demo-a'));runtime.close();});
test('concurrent selection cannot resurrect old company',async()=>{const f=fixture(),runtime=createOriginalRuntimeSession(f);const a=runtime.select('demo-a'),b=runtime.select('demo-b');await assert.rejects(a,/cancelled/);assert.equal((await b).context.tenantId,'demo-b');runtime.close();});

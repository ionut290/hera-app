import test from 'node:test';import assert from 'node:assert/strict';import {createLegacyFirebaseNamespace} from './legacy-firebase-namespace.mjs';
function fixture(){
 let active=true,observer;const projectId='varga-cantieri-v2-staging';
 const raw={uid:'u',email:'u@example.test',getIdTokenResult:async()=>({claims:{email_verified:true}}),auth:{secret:true}};
 const auth={app:{options:{projectId},firestore(){throw Error('raw');}},currentUser:raw,onAuthStateChanged(fn){observer=fn;return()=>{};},signOut:async()=>{active=false;}};
 const firestore={collection:n=>({path:'tenants/demo-a/'+n})};
 const prepared={projectId,auth,session:{context:{uid:'u',tenantId:'demo-a'},firestore},isActive:()=>active};
 return {prepared,raw,firestore,close(){active=false;},emit:u=>observer(u)};
}
test('direct global Firebase callers receive the same tenant scope, not a raw SDK app',()=>{
 const f=fixture(),sentinel={serverTimestamp:()=>42},sdk=createLegacyFirebaseNamespace({prepared:f.prepared,statics:{FieldValue:sentinel}});
 assert.equal(sdk.firestore(),f.firestore);assert.equal(sdk.app().firestore(),f.firestore);assert.equal(sdk.firestore.FieldValue.serverTimestamp(),42);
 assert.equal(sdk.auth().app,sdk.app());assert.notEqual(sdk.auth().currentUser,f.raw);assert.equal(sdk.auth().currentUser.auth,undefined);
 for(const method of ['storage','functions','messaging','initializeApp'])assert.throws(()=>sdk[method]());
 assert.throws(()=>sdk.app('production'));assert.throws(()=>sdk.app().functions('europe-west1'));
});
test('namespace cannot reopen a company after switch and observers cannot reveal a new user',()=>{
 const f=fixture(),sdk=createLegacyFirebaseNamespace({prepared:f.prepared});const auth=sdk.auth(),received=[];
 auth.onAuthStateChanged(u=>received.push(u));f.emit(f.raw);assert.equal(received[0].uid,'u');f.emit({uid:'other'});assert.equal(received[1],null);
 f.close();assert.equal(auth.currentUser,null);assert.throws(()=>sdk.firestore());assert.throws(()=>sdk.app());assert.throws(()=>sdk.auth());assert.throws(()=>auth.setPersistence());
});
test('production and incorrect user are rejected',()=>{const f=fixture();f.prepared.auth.app.options.projectId='hera-app-6cd2b';assert.throws(()=>createLegacyFirebaseNamespace({prepared:f.prepared}));});

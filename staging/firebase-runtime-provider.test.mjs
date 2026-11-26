import test from 'node:test';import assert from 'node:assert/strict';import {connectLocalFirebase,stagingProjectId} from './firebase-runtime-provider.mjs';
function fixture(){
 const calls=[];const auth={useEmulator:(...x)=>calls.push(['auth',...x]),setPersistence:async x=>calls.push(['persistence',x])},db={useEmulator:(...x)=>calls.push(['db',...x])};
 const firebase={initializeApp:c=>{calls.push(['app',c]);return {auth:()=>auth,firestore:()=>db};},auth:{Auth:{Persistence:{NONE:'none'}}},firestore:{FieldValue:{},Timestamp:{}}};return {firebase,calls,auth,db};
}
test('provider only initializes verified staging emulator endpoints',async()=>{const f=fixture(),r=await connectLocalFirebase({firebase:f.firebase,location:{hostname:'127.0.0.1',protocol:'http:'}});assert.equal(r.auth,f.auth);assert.equal(r.db,f.db);assert.equal(f.calls[0][1].projectId,stagingProjectId);assert.deepEqual(f.calls.slice(1),[['auth','http://127.0.0.1:9099',{disableWarnings:true}],['db','127.0.0.1',8080],['persistence','none']]);});
test('provider refuses external sites before SDK initialization',async()=>{const f=fixture();await assert.rejects(connectLocalFirebase({firebase:f.firebase,location:{hostname:'example.com',protocol:'https:'}}));assert.equal(f.calls.length,0);});

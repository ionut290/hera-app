import test from 'node:test';
import assert from 'node:assert/strict';
import {saveFattoSettings} from './fatto-settings-store.mjs';
function fixture({mode='configurable-demo',role='admin',version=0,project='varga-cantieri-v2-staging',verified=true}={}){
 const writes=[];
 const data={'tenants/demo-a':{demo:true,status:'active',fattoMode:mode},'tenants/demo-a/members/u':{status:'active',role},'users/u/memberships/demo-a':{status:'active',role,tenantId:'demo-a'}};
 if(version)data['tenants/demo-a/fattoSettings/workflow']={version};
 const sdk={doc:(_, ...p)=>p.join('/'),serverTimestamp:()=>123,runTransaction:async(_,fn)=>fn({get:async p=>({data:()=>data[p],exists:()=>!!data[p]}),set:(p,d)=>writes.push([p,d])})};
 return {writes,args:{db:{app:{options:{projectId:project}}},user:{uid:'u',getIdTokenResult:async()=>({claims:{email_verified:verified}})},tenantId:'demo-a',actions:['saveCompletion'],expectedVersion:0,sdk}};
}
test('creates tenant scoped versioned policy',async()=>{const f=fixture();await saveFattoSettings(f.args);assert.deepEqual(f.writes,[['tenants/demo-a/fattoSettings/workflow',{version:1,actions:['saveCompletion'],updatedAt:123,updatedBy:'u'}]]);});
test('updates matching version',async()=>{const f=fixture({version:3});await saveFattoSettings({...f.args,expectedVersion:3});assert.equal(f.writes[0][1].version,4);});
for(const [name,opts] of Object.entries({protected:{mode:'immutable'},employee:{role:'employee'},stale:{version:2},production:{project:'hera-app-6cd2b'},unverified:{verified:false}}))test('reject '+name,async()=>{const f=fixture(opts);await assert.rejects(saveFattoSettings(f.args));assert.equal(f.writes.length,0);});

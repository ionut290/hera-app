import test from 'node:test';
import assert from 'node:assert/strict';
import {saveCoboDemo} from './cobo-transaction.mjs';
function fixture({role='admin',project='varga-cantieri-v2-staging',commessa=false,plant=false}={}){
 const writes=[];
 const data={
 'tenants/demo-a':{demo:true,status:'active',coboWriteEnabled:true},
 'tenants/demo-a/members/u':{status:'active',role},
 'users/u/memberships/demo-a':{status:'active',role,tenantId:'demo-a'},
 ...(commessa?{'tenants/demo-a/commesse/sfalcio-cobo':{nome:'Sfalcio COBO'}}:{}),
 ...(plant?{'tenants/demo-a/commesse/sfalcio-cobo/impianti/cobo-parco-test':{done:false}}:{})
 };
 const sdk={doc:(_, ...p)=>p.join('/'),serverTimestamp:()=>({server:true}),runTransaction:async(_,fn)=>fn({get:async p=>({exists:()=>!!data[p],data:()=>data[p]}),set:(p,v,o)=>writes.push({path:p,value:v,options:o})})};
 const args={db:{app:{options:{projectId:project}}},user:{uid:'u',email:'u@example.test',displayName:'User',getIdTokenResult:async()=>({claims:{email_verified:true}})},tenantId:'demo-a',payload:{id:'cobo-parco-test',commessaId:'sfalcio-cobo',sfalcioCobo:true,denominazione:'Parco demo'},sdk};
 return {args,writes};
}
test('admin atomically initializes commessa and plant',async()=>{const f=fixture();const result=await saveCoboDemo(f.args);assert.equal(result.created,true);assert.equal(f.writes.length,2);assert.equal(f.writes[0].path,'tenants/demo-a/commesse/sfalcio-cobo');assert.equal(f.writes[1].path,'tenants/demo-a/commesse/sfalcio-cobo/impianti/cobo-parco-test');});
test('employee can write plant when commessa exists',async()=>{const f=fixture({role:'employee',commessa:true});const result=await saveCoboDemo(f.args);assert.equal(result.created,true);assert.equal(f.writes.length,1);});
test('employee cannot initialize commessa',async()=>{const f=fixture({role:'employee'});await assert.rejects(saveCoboDemo(f.args));assert.equal(f.writes.length,0);});
test('reject production project',async()=>{const f=fixture({project:'hera-production'});await assert.rejects(saveCoboDemo(f.args));});
test('reject unexpected payload field',async()=>{const f=fixture();f.args.payload.done=true;await assert.rejects(saveCoboDemo(f.args));});
test('reject unverified user',async()=>{const f=fixture();f.args.user.getIdTokenResult=async()=>({claims:{email_verified:false}});await assert.rejects(saveCoboDemo(f.args));});
test('preserves existing completion state on merge',async()=>{const f=fixture({commessa:true,plant:true});await saveCoboDemo(f.args);assert.equal(f.writes[0].value.done,undefined);assert.equal(f.writes[0].options.merge,true);});

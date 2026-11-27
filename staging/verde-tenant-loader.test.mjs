import test from 'node:test';
import assert from 'node:assert/strict';
import {loadTenantVerde} from './verde-tenant-loader.mjs';
function fixture({role='admin',membershipRole=role,status='active',project='varga-cantieri-v2-staging',verified=true}={}){
 const calls=[];
 const db={app:{options:{projectId:project}}};
 const user={uid:'u1',getIdTokenResult:async force=>{assert.equal(force,true);return {claims:{email_verified:verified}}}};
 const snap=data=>({exists:()=>true,data:()=>data});
 const sdk={doc:(_,...p)=>p.join('/'),collection:(_,...p)=>p.join('/'),getDocFromServer:async path=>{
 calls.push(path);
 if(path==='tenants/demo-a')return snap({demo:true,status:'active'});
 if(path==='tenants/demo-a/members/u1')return snap({role,status});
 if(path==='users/u1/memberships/demo-a')return snap({tenantId:'demo-a',role:membershipRole,status:'active'});
 throw Error('Unexpected path');
 },getDocsFromServer:async path=>{calls.push(path);return {docs:[]}}};
 return {args:{db,user,tenantId:'demo-a',sdk},calls};
}
test('authorized empty tenant',async()=>{const {args,calls}=fixture();assert.deepEqual(await loadTenantVerde(args),[]);assert.equal(calls.at(-1),'tenants/demo-a/verdeLevatoRecords');});
for(const options of [{role:'unknown'},{membershipRole:'employee'},{status:'suspended'},{verified:false},{project:'production'}]){
 test('deny '+JSON.stringify(options),async()=>{const {args,calls}=fixture(options);await assert.rejects(loadTenantVerde(args));assert.ok(!calls.some(p=>p.includes('verdeLevatoRecords')));});
}

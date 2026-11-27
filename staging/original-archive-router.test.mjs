import test from 'node:test';
import assert from 'node:assert/strict';
import {createOriginalArchiveRouter} from './original-archive-router.mjs';
function fixture(tenantId='demo-a'){
 const docs=new Map([[ 'tenants/'+tenantId,{demo:true,status:'active'}],['tenants/'+tenantId+'/members/u',{role:'admin',status:'active'}],['users/u/memberships/'+tenantId,{tenantId,role:'admin',status:'active'}]]);
 const db={app:{options:{projectId:'varga-cantieri-v2-staging'}}};
 const user={uid:'u',getIdTokenResult:async()=>({claims:{email_verified:true}})};
 const sdk={doc:(_, ...p)=>p.join('/'),getDocFromServer:async p=>({exists:()=>docs.has(p),data:()=>docs.get(p)})};
 return {db,user,tenantId,sdk,docs,getCurrentUser:()=>user};
}
test('original archives preserve names within tenant',async()=>{const r=await createOriginalArchiveRouter(fixture());for(const name of ['commesse','neve_commesse','globalCommesse','personale','mezzi','neve_personale','neve_mezzi','appConfig','oreLocks','oreReports','calendarEvents','verdeLevatoRecords'])assert.equal(r.path(name),'tenants/demo-a/'+name);});
test('same archive names separate companies',async()=>{const a=await createOriginalArchiveRouter(fixture()),b=await createOriginalArchiveRouter(fixture('demo-b'));assert.notEqual(a.path('commesse'),b.path('commesse'));});
test('unreviewed, sensitive and injected collections denied',async()=>{const r=await createOriginalArchiveRouter(fixture());for(const n of ['platformUsers','users','driveAdminSecret','sheetExportQueue','tenants/demo-b/commesse','../commesse','unknown'])assert.throws(()=>r.path(n));});
test('closed scope cannot be reused',async()=>{const r=await createOriginalArchiveRouter(fixture());r.close();assert.throws(()=>r.path('commesse'));});
test('production rejected',async()=>{const f=fixture();f.db.app.options.projectId='hera-app-6cd2b';await assert.rejects(createOriginalArchiveRouter(f));});
test('unverified account rejected',async()=>{const f=fixture();f.user.getIdTokenResult=async()=>({claims:{email_verified:false}});await assert.rejects(createOriginalArchiveRouter(f));});
test('missing membership denied',async()=>{const f=fixture();f.docs.delete('users/u/memberships/demo-a');await assert.rejects(createOriginalArchiveRouter(f));});
test('mismatched role denied',async()=>{const f=fixture();f.docs.get('users/u/memberships/demo-a').role='employee';await assert.rejects(createOriginalArchiveRouter(f));});
test('suspended tenant denied',async()=>{const f=fixture();f.docs.get('tenants/demo-a').status='suspended';await assert.rejects(createOriginalArchiveRouter(f));});
test('review account denied',async()=>{const f=fixture();f.user.getIdTokenResult=async()=>({claims:{email_verified:true,apple_review:true}});await assert.rejects(createOriginalArchiveRouter(f));});

test('logout or user change invalidates archive router',async()=>{const f=fixture();let current=f.user;f.getCurrentUser=()=>current;const r=await createOriginalArchiveRouter(f);current=null;assert.throws(()=>r.path('commesse'),/Context closed/);});
test('session change during server verification denies access',async()=>{const f=fixture();let current=f.user;f.getCurrentUser=()=>current;const read=f.sdk.getDocFromServer;f.sdk.getDocFromServer=async p=>{current={uid:'other'};return read(p);};await assert.rejects(createOriginalArchiveRouter(f),/Access denied/);});

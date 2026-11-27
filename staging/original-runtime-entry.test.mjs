import test from 'node:test';import assert from 'node:assert/strict';
import {createEphemeralRuntimeStorage,createOriginalRuntimeEntry} from './original-runtime-entry.mjs';
function fixture(role='employee'){
 const app={options:{projectId:'varga-cantieri-v2-staging'}},user={uid:'u',email:'owner@example.test',getIdTokenResult:async()=>({claims:{email_verified:true,superAdmin:true}})};
 const records=new Map(),listeners=new Set();
 for(const id of ['demo-a','demo-b']){
  records.set('tenants/'+id,{demo:true,status:'active',name:id});
  records.set('tenants/'+id+'/members/u',{role,status:'active'});
  records.set('users/u/memberships/'+id,{role,status:'active',tenantId:id});
 }
 const auth={app,currentUser:user,onAuthStateChanged(fn){listeners.add(fn);return()=>listeners.delete(fn);}};
 const db={app,doc:path=>({path,get:async opt=>{assert.equal(opt.source,'server');return {exists:records.has(path),data:()=>records.get(path)};}}),collection:path=>({path})};
 const entry=createOriginalRuntimeEntry({auth,db});
 return {entry,auth,db,records,user,logout(){auth.currentUser=null;for(const fn of [...listeners])fn(null);}};
}
test('legacy cached admin or global superAdmin claim does not grant company administration',async()=>{
 const f=fixture(),p=await f.entry.enter('demo-a');
 p.localStorage.setItem('heraPersistedUserSession',JSON.stringify({isAdmin:true}));
 const verified=await p.approval.verify({...f.user,role:'admin',isAdmin:true});
 assert.equal(verified.allowed,true);assert.equal(verified.profile.isAdmin,false);assert.equal(verified.profile.tenantRole,'employee');assert.equal(verified.profile.tenantId,'demo-a');f.entry.close();
});
test('company cache does not inherit keys, closes on switch, logout and suspension',async()=>{
 const f=fixture('admin'),a=await f.entry.enter('demo-a');a.localStorage.setItem('heraCommesseCache','private-a');a.sessionStorage.setItem('queue','a');
 const b=await f.entry.enter('demo-b');assert.throws(()=>a.localStorage.getItem('heraCommesseCache'));assert.throws(()=>a.sessionStorage.getItem('queue'));assert.equal(b.localStorage.getItem('heraCommesseCache'),null);assert.equal((await b.approval.verify(f.user)).profile.isAdmin,true);
 f.records.get('tenants/demo-b/members/u').status='suspended';assert.equal((await b.approval.verify(f.user)).allowed,false);assert.throws(()=>b.sessionStorage.setItem('queue','x'));
 const c=await f.entry.enter('demo-a');f.logout();assert.equal(c.isActive(),false);assert.throws(()=>c.localStorage.length);f.entry.close();
});
test('role change requires a fresh company entry instead of silently retaining old authority',async()=>{
 const f=fixture('admin'),p=await f.entry.enter('demo-a');
 f.records.get('tenants/demo-a/members/u').role='employee';f.records.get('users/u/memberships/demo-a').role='employee';
 assert.equal((await p.approval.verify(f.user)).allowed,false);assert.equal(p.isActive(),false);
 const fresh=await f.entry.enter('demo-a');assert.equal((await fresh.approval.verify(f.user)).profile.isAdmin,false);f.entry.close();
});
test('storage has bounded quota, removes only its own in-memory keys, and cannot reopen',()=>{
 const x=createEphemeralRuntimeStorage({maxBytes:12}),s=x.storage;s.setItem('a','123');assert.equal(s.key(0),'a');assert.throws(()=>s.setItem('b','123'),/quota/);assert.equal(s.getItem('a'),'123');s.removeItem('a');s.setItem('b','123');assert.equal(s.length,1);s.clear();assert.equal(s.length,0);x.close();assert.throws(()=>s.setItem('a','b'));
});

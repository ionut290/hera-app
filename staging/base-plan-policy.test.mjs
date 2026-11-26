import test from 'node:test';
import assert from 'node:assert/strict';
import {BASE_PLAN,quotaAfterCreate,memberQuotaKind,validateLimits,decideExpansion,validateExpansion} from './base-plan-policy.mjs';
const usage={commesse:0,impianti:0,admin:0,operatori:0};
for(const kind of Object.keys(BASE_PLAN)){
 test('base limit: '+kind,()=>{
  const after=quotaAfterCreate({limits:BASE_PLAN,usage,kind,amount:BASE_PLAN[kind]});
  assert.equal(after[kind],BASE_PLAN[kind]);
  assert.throws(()=>quotaAfterCreate({limits:BASE_PLAN,usage:after,kind}),/SUPER_ADMIN_APPROVAL_REQUIRED/);
  assert.equal(usage[kind],0);
 });
}
test('roles cannot bypass operator limits',()=>{
 assert.equal(memberQuotaKind('admin'),'admin');
 assert.equal(memberQuotaKind('employee'),'operatori');
 assert.equal(memberQuotaKind('supervisor'),'operatori');
 assert.throws(()=>memberQuotaKind('superAdmin'),/INVALID_COMPANY_ROLE/);
});
test('invalid inputs fail closed',()=>{
 assert.throws(()=>validateLimits({...BASE_PLAN,admin:0}),/INVALID_LIMITS/);
 assert.throws(()=>quotaAfterCreate({limits:BASE_PLAN,usage,kind:'operatori',amount:0}),/INVALID_INCREMENT/);
 assert.throws(()=>quotaAfterCreate({limits:BASE_PLAN,usage:{...usage,operatori:-1},kind:'operatori'}),/INVALID_USAGE/);
 assert.throws(()=>validateExpansion({currentLimits:BASE_PLAN,requestedLimits:BASE_PLAN}),/NO_EXPANSION_REQUESTED/);
});
const actor={uid:'owner',superAdmin:true,email_verified:true};
const request={tenantId:'demo-a',status:'pending',baseLimits:BASE_PLAN,requestedLimits:{...BASE_PLAN,impianti:20}};
const args={actor,tenantId:'demo-a',request,currentLimits:BASE_PLAN,decision:'approved'};
test('only verified superadmin can approve same-company pending expansion',()=>{
 const result=decideExpansion(args);
 assert.equal(result.limits.impianti,20);
 assert.equal(result.requestPatch.reviewedBy,'owner');
 assert.throws(()=>decideExpansion({...args,actor:{uid:'admin'}}),/SUPER_ADMIN_REQUIRED/);
 assert.throws(()=>decideExpansion({...args,actor:{...actor,email_verified:false}}),/SUPER_ADMIN_REQUIRED/);
 assert.throws(()=>decideExpansion({...args,tenantId:'demo-b'}),/INVALID_PENDING_REQUEST/);
 assert.throws(()=>decideExpansion({...args,request:{...request,status:'approved'}}),/INVALID_PENDING_REQUEST/);
});
test('rejection never expands and stale approval fails',()=>{
 assert.deepEqual(decideExpansion({...args,decision:'rejected'}).limits,BASE_PLAN);
 assert.throws(()=>decideExpansion({...args,currentLimits:{...BASE_PLAN,admin:2}}),/STALE_REQUEST/);
});

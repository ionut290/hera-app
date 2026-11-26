import test from 'node:test';import assert from 'node:assert/strict';import {readFileSync} from 'node:fs';import {Script,createContext} from 'node:vm';
import {buildIntegratedRuntimeCandidate,integratedStartupGate,integratedAdminGate} from './original-runtime-source.mjs';
import {createEphemeralRuntimeStorage} from './original-runtime-entry.mjs';
const original=readFileSync(new URL('../app.js',import.meta.url),'utf8');
test('full original source compiles with company startup, role and cache binding',()=>{
 const candidate=buildIntegratedRuntimeCandidate(original);
 assert.doesNotThrow(()=>new Script(candidate));
 assert(candidate.includes(integratedStartupGate));
 assert(candidate.includes(integratedAdminGate));
 assert(candidate.includes('const BUILT_IN_SUPER_ADMIN_EMAILS = [];'));
 assert(!candidate.includes('const hasSavedApproval = isPersistedApprovalValid(savedSession, user);'));
 // Protected operations are not rewritten by this integration.
 for(const name of ['getImpiantoDocIds','getMarkerClass','updateImpiantoLocalState']){
  const start=original.indexOf('function '+name+'('),end=original.indexOf('\n}\n',start);
  assert(start>=0&&end>start);assert(candidate.includes(original.slice(start,end+2)));
 }
});
async function runGate({role='employee',allowed=true,active=true,uid='u',userUid='u',throwVerify=false}={}){
 const effects=[],local=createEphemeralRuntimeStorage(),session=createEphemeralRuntimeStorage();
 const prepared={session:{context:{uid,tenantId:'demo-a',role}},isActive:()=>active,approval:{verify:async()=>{effects.push('verify');if(throwVerify)throw Error('offline');return {allowed,profile:{tenantRole:role,tenantId:'demo-a'}};}},localStorage:local.storage,sessionStorage:session.storage};
 const context=createContext({window:{__VARGA_ISOLATED_RUNTIME_CANDIDATE__:prepared},user:{uid:userUid},currentUser:{uid:userUid},loggedIn:true,savePersistedSession:async(_u,p)=>effects.push('save:'+p.tenantRole),setAuthenticationGateState:()=>effects.push('required'),hideStartupLoading:()=>effects.push('hide')});
 const fn=new Script('(async()=>{'+integratedStartupGate+'} effectsMarker(); return "ok";})');
 context.effectsMarker=()=>effects.push('continue');
 await fn.runInContext(context)();
 const admin=new Script(integratedAdminGate+'\ncanManageData()').runInContext(context);
 return {effects,admin};
}
test('actual transformed original startup gate accepts verified tenant roles only',async()=>{
 assert.deepEqual(await runGate({role:'employee'}),{effects:['verify','save:employee','continue'],admin:false});
 assert.deepEqual(await runGate({role:'admin'}),{effects:['verify','save:admin','continue'],admin:true});
});
test('actual transformed gate stops before operational startup for stale, denied or offline users',async()=>{
 for(const opts of [{active:false},{allowed:false},{uid:'other'},{throwVerify:true}]){
  const r=await runGate({...opts,role:'admin'});assert(r.effects.includes('required'));assert(!r.effects.includes('continue'));assert.equal(r.admin,false);
 }
});

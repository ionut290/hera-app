import test from 'node:test';
import assert from 'node:assert/strict';
import {validateFattoPolicy} from './fatto-policy.mjs';
const tenant={id:'demo-001',demo:true,status:'active',fattoMode:'configurable-demo',sheetConfigured:true,whatsappConfigured:true};
const context={uid:'admin-a',tenantId:tenant.id,role:'admin'};
const policy={version:1,actions:['saveCompletion','moveToCompleted','openWhatsApp','exportSheet']};
const valid={tenant,context,policy};
test('valid workflow immutable',()=>{const p=validateFattoPolicy(valid);assert.deepEqual(p.actions,policy.actions);assert.ok(Object.isFrozen(p.actions));});
for(const [name,change] of Object.entries({
 employee:{context:{...context,role:'employee'}},
 supervisor:{context:{...context,role:'supervisor'}},
 crossTenant:{context:{...context,tenantId:'demo-002'}},
 missingUid:{context:{...context,uid:''}},
 protected:{tenant:{...tenant,fattoMode:'immutable'}},
 unknown:{tenant:{...tenant,fattoMode:undefined}},
 real:{tenant:{...tenant,demo:false}},
 suspended:{tenant:{...tenant,status:'suspended'}},
 whatsapp:{tenant:{...tenant,whatsappConfigured:false}},
 sheet:{tenant:{...tenant,sheetConfigured:false}},
 duplicate:{policy:{...policy,actions:['saveCompletion','saveCompletion']}},
 arbitrary:{policy:{...policy,actions:['saveCompletion','eval']}},
 wrongOrder:{policy:{...policy,actions:['openWhatsApp','saveCompletion']}},
 empty:{policy:{...policy,actions:[]}},
 version:{policy:{...policy,version:0}}
}))test('reject '+name,()=>assert.throws(()=>validateFattoPolicy({...valid,...change})));
test('optional integrations may be disabled',()=>assert.deepEqual(validateFattoPolicy({...valid,policy:{version:2,actions:['saveCompletion']}}).actions,['saveCompletion']));

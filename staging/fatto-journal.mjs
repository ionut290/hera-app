import {startExecution,transitionExecution} from './fatto-execution-state.mjs';
// Serialization only. No Firestore write permission or external action is enabled.
export function journalPath(context,operationId){
 if(!context?.uid||!['admin','supervisor','employee'].includes(context.role)||typeof context.tenantId!=='string'||!/^[a-z0-9-]{3,64}$/.test(context.tenantId))throw Error('Invalid context');
 if(typeof operationId!=='string'||!/^[a-zA-Z0-9_-]{1,128}$/.test(operationId))throw Error('Invalid operation ID');
 return ['tenants',context.tenantId,'fattoExecutions',operationId];
}
export function restoreJournal(plan,record){
 if(!record||record.tenantId!==plan.tenantId||record.operationId!==plan.operationId||record.siteId!==plan.siteId||record.version!==plan.version||!Array.isArray(record.steps)||record.steps.length!==plan.steps.length)throw Error('Journal mismatch');
 let state=startExecution(plan);
 for(let i=0;i<plan.steps.length;i++){
 const expected=plan.steps[i],saved=record.steps[i];
 if(saved.key!==expected.key||saved.action!==expected.action)throw Error('Step mismatch');
 if(saved.status==='pending')continue;
 state=transitionExecution(state,expected.key,'start');
 const event={done:'success',failed:'failure',running:'uncertain',uncertain:'uncertain'}[saved.status];
 if(!event)throw Error('Invalid status');
 state=transitionExecution(state,expected.key,event);
 }
 return state;
}
export function serializeJournal(state){
 const {tenantId,operationId,siteId,version}=state.plan;
 return {tenantId,operationId,siteId,version,steps:state.steps.map(({key,action,status})=>({key,action,status}))};
}

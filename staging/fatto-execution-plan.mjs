import {validateFattoPolicy} from './fatto-policy.mjs';
// Planning only. Trusted server-side authorization and durable execution are required before use.
export function planFatto({tenant,context,policy,siteId,operationId}){
 if(!context?.uid||!['admin','supervisor','employee'].includes(context.role))throw Error('Access denied');
 for(const id of [siteId,operationId])if(typeof id!=='string'||!id.trim()||id.includes('/')||id==='.'||id==='..')throw Error('Invalid ID');
 const validated=validateFattoPolicy({tenant,context:{...context,role:'admin'},policy});
 return Object.freeze({tenantId:validated.tenantId,siteId,operationId,version:validated.version,steps:Object.freeze(validated.actions.map((action,index)=>Object.freeze({action,index,key:JSON.stringify([validated.tenantId,siteId,operationId,index,action])})))});
}

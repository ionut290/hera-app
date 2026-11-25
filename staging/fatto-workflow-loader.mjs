import {resolveTenantContext} from './tenant-context.mjs';
import {planFatto} from './fatto-execution-plan.mjs';
import {startExecution} from './fatto-execution-state.mjs';
// Read-only planning; no action execution. Employee policy read rules must be approved before activation.
export async function loadFattoWorkflow({db,user,tenantId,siteId,operationId,sdk}){
 if(db?.app?.options?.projectId!=='varga-cantieri-v2-staging')throw Error('Staging required');
 if(!user?.uid||typeof tenantId!=='string'||!/^[a-z0-9-]{3,64}$/.test(tenantId))throw Error('Invalid request');
 const token=await user.getIdTokenResult(true);
 if(token.claims?.email_verified!==true||token.claims?.apple_review===true)throw Error('Access denied');
 const {doc,getDocFromServer}=sdk;
 const read=async(...p)=>{const s=await getDocFromServer(doc(db,...p));return s.exists()?s.data():null;};
 const [tenant,member,membership]=await Promise.all([read('tenants',tenantId),read('tenants',tenantId,'members',user.uid),read('users',user.uid,'memberships',tenantId)]);
 const context=resolveTenantContext({uid:user.uid,claims:token.claims,tenantId,tenant,member,membership});
 if(!context||tenant.fattoMode!=='configurable-demo')throw Error('Protected or unauthorized workflow');
 const policy=await read('tenants',tenantId,'fattoSettings','workflow');
 if(!policy)throw Error('Configuration required');
 return startExecution(planFatto({tenant:{...tenant,id:tenantId},context,policy,siteId,operationId}));
}

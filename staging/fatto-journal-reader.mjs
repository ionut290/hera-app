import {resolveTenantContext} from './tenant-context.mjs';
import {journalPath,restoreJournal} from './fatto-journal.mjs';
// Read-only adapter. Journal rules and trusted writer are not yet deployed.
export async function readFattoJournal({db,user,plan,sdk}){
 if(db?.app?.options?.projectId!=='varga-cantieri-v2-staging')throw Error('Staging required');
 if(!user?.uid||typeof plan?.tenantId!=='string'||!/^[a-z0-9-]{3,64}$/.test(plan.tenantId))throw Error('Invalid request');
 const tenantId=plan.tenantId;
 const token=await user.getIdTokenResult(true);
 if(token.claims?.email_verified!==true||token.claims?.apple_review===true)throw Error('Access denied');
 const {doc,getDocFromServer}=sdk;
 const read=async(...path)=>{const snap=await getDocFromServer(doc(db,...path));return snap.exists()?snap.data():null;};
 const [tenant,member,membership]=await Promise.all([read('tenants',tenantId),read('tenants',tenantId,'members',user.uid),read('users',user.uid,'memberships',tenantId)]);
 const context=resolveTenantContext({uid:user.uid,claims:token.claims,tenantId,tenant,member,membership});
 if(!context||tenant.fattoMode!=='configurable-demo')throw Error('Access denied');
 const record=await read(...journalPath(context,plan.operationId));
 if(!record)return null;
 return restoreJournal(plan,record);
}

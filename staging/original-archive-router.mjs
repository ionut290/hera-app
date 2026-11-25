// Staging adapter only; not installed in the legacy runtime.
import {resolveTenantContext} from './tenant-context.mjs';
const operational = new Set(['commesse','neve_commesse','globalCommesse','squadreCommesse','squadreStorico','neve_squadre','neve_squadre_storico','personale','neve_personale','mezzi','neve_mezzi','oreReports','neve_ore','oreApprovalRequests','neve_ore_richieste','programmazioni','calendarEvents','commessaResources','verdeLevatoRecords','verdeLevatoCommesse','verdeLevatoConfig']);
const allowedRoles=new Set(['admin','supervisor','employee']);
export async function createOriginalArchiveRouter({db,user,tenantId,sdk}){
 if(db?.app?.options?.projectId!=='varga-cantieri-v2-staging')throw Error('Staging required');
 if(!user?.uid||typeof tenantId!=='string'||!/^[a-z0-9-]{3,64}$/.test(tenantId))throw Error('Invalid context');
 const uid=user.uid;
 const token=await user.getIdTokenResult(true);
 if(token.claims?.email_verified!==true||token.claims?.apple_review===true)throw Error('Access denied');
 const read=async(...path)=>{const s=await sdk.getDocFromServer(sdk.doc(db,...path));return s.exists()?s.data():null;};
 const [tenant,member,membership]=await Promise.all([read('tenants',tenantId),read('tenants',tenantId,'members',uid),read('users',uid,'memberships',tenantId)]);
 const context=resolveTenantContext({uid,claims:token.claims,tenantId,tenant,member,membership});
 if(!context||!allowedRoles.has(context.role))throw Error('Access denied');
 let closed=false;
 function path(name){
  if(closed||user.uid!==uid)throw Error('Context closed');
  if(!operational.has(name))throw Error('Collection not approved');
  return 'tenants/'+context.tenantId+'/'+name;
 }
 return Object.freeze({projectId:db.app.options.projectId,context,path,close(){closed=true;}});
}

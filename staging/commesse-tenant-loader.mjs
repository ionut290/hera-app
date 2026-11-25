import {resolveTenantContext} from './tenant-context.mjs';
import {createCommesseStore} from './commesse-store.mjs';
export async function loadTenantCommesseStore({db,user,tenantId,sdk}){
 if(db?.app?.options?.projectId!=='varga-cantieri-v2-staging')throw Error('Staging required');
 if(!user?.uid||typeof tenantId!=='string'||!/^[a-z0-9-]{3,64}$/.test(tenantId))throw Error('Invalid request');
 const token=await user.getIdTokenResult(true);
 if(token.claims?.email_verified!==true||token.claims?.apple_review===true)throw Error('Access denied');
 const read=async(...path)=>{const snap=await sdk.getDocFromServer(sdk.doc(db,...path));return snap.exists()?snap.data():null;};
 const [tenant,member,membership]=await Promise.all([read('tenants',tenantId),read('tenants',tenantId,'members',user.uid),read('users',user.uid,'memberships',tenantId)]);
 const context=resolveTenantContext({uid:user.uid,claims:token.claims,tenantId,tenant,member,membership});
 if(!context)throw Error('Access denied');
 return createCommesseStore(db,context,sdk);
}

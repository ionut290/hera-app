import {resolveTenantContext} from './tenant-context.mjs';
import {createModularVerdeStore} from './verde-modular-store.mjs';
import {projectVerdeRecords} from './verde-record-projection.mjs';
export async function loadTenantVerde({db,user,tenantId,sdk}) {
  if (db?.app?.options?.projectId !== 'varga-cantieri-v2-staging') throw new Error('Staging Firebase required');
  if (!user?.uid || typeof tenantId !== 'string' || !/^[a-z0-9-]{3,64}$/.test(tenantId)) throw new Error('Invalid tenant request');
  const token=await user.getIdTokenResult(true);
  const {doc,getDocFromServer}=sdk;
  const [tenant,member,membership]=await Promise.all([
    getDocFromServer(doc(db,'tenants',tenantId)),
    getDocFromServer(doc(db,'tenants',tenantId,'members',user.uid)),
    getDocFromServer(doc(db,'users',user.uid,'memberships',tenantId))
  ]);
  const data=s=>s.exists()?s.data():null;
  const context=resolveTenantContext({uid:user.uid,claims:token.claims,tenantId,tenant:data(tenant),member:data(member),membership:data(membership)});
  if (!context) throw new Error('Tenant access denied');
  const store=createModularVerdeStore(db,context,sdk);
  return projectVerdeRecords(await store.collection('verdeLevatoRecords').get());
}

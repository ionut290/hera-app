import {validateFattoPolicy} from './fatto-policy.mjs';
// Not wired to UI until full deployed rules are verified.
export async function saveFattoSettings({db,user,tenantId,actions,expectedVersion,sdk}){
 if(db?.app?.options?.projectId!=='varga-cantieri-v2-staging')throw Error('Staging required');
 if(!user?.uid || !/^[a-z0-9-]{3,64}$/.test(tenantId)||!Number.isSafeInteger(expectedVersion)||expectedVersion<0)throw Error('Invalid request');
 const token=await user.getIdTokenResult(true);
 if(token.claims?.email_verified!==true||token.claims?.apple_review===true)throw Error('Access denied');
 const {doc,runTransaction,serverTimestamp}=sdk;
 return runTransaction(db,async tx=>{
 const tenantRef=doc(db,'tenants',tenantId);
 const memberRef=doc(db,'tenants',tenantId,'members',user.uid);
 const membershipRef=doc(db,'users',user.uid,'memberships',tenantId);
 const settingRef=doc(db,'tenants',tenantId,'fattoSettings','workflow');
 const [t,m,u]=await Promise.all([tx.get(tenantRef),tx.get(memberRef),tx.get(membershipRef)]);
 const tenant=t.data(),member=m.data(),membership=u.data();
 if(!tenant||member?.status!=='active'||member.role!=='admin'||membership?.status!=='active'||membership.role!=='admin'||membership.tenantId!==tenantId)throw Error('Access denied');
 const validated=validateFattoPolicy({tenant:{...tenant,id:tenantId},context:{uid:user.uid,tenantId,role:member.role},policy:{version:expectedVersion+1,actions}});
 const current=await tx.get(settingRef);
 if((current.exists()?current.data().version:0)!==expectedVersion)throw Error('Configuration changed; reload');
 tx.set(settingRef,{version:validated.version,actions:[...validated.actions],updatedAt:serverTimestamp(),updatedBy:user.uid});
 return validated;
 });
}

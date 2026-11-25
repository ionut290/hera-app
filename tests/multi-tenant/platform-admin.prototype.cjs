'use strict';
// Isolated prototype; never exported by the live app or deployed.
// Auth must come from a server-side verified, non-revoked Firebase ID token.
const {randomUUID}=require('node:crypto');
function fail(code){throw Object.assign(new Error(code),{code});}
function requireAdmin(auth){
  if(!auth||auth.email_verified!==true||auth.apple_review===true||auth.disabled===true)fail('unauthenticated');
  if(auth.platform_admin!==true)fail('forbidden');
  return safeId(auth.uid);
}
function safeId(value){
  if(typeof value!=='string'||!(/^[A-Za-z0-9_-]{1,128}$/).test(value)||/^__.*__$/.test(value))fail('invalid-id');
  return value;
}
function publicCompany(id,data){
  return {id,name:String(data.name||''),status:data.status,createdAt:data.createdAt??null};
}
// store.listTenants({limit,after}) must return at most limit+1 entries sorted by document ID.
// store.runTransaction contract is the same as tenant-service.prototype.cjs.
function makePlatformAdminService(store,clock=Date.now){
  async function listCompanies(auth,{limit=25,after=null}={}){
    requireAdmin(auth);
    if(!Number.isInteger(limit)||limit<1||limit>100)fail('invalid-limit');
    if(after!==null)safeId(after);
    const rows=await store.listTenants({limit,after});
    return {companies:rows.slice(0,limit).map(row=>publicCompany(row.id,row.data)),
      nextCursor:rows.length>limit?rows[limit-1].id:null};
  }
  async function setCompanyStatus(auth,{tenantId,status,reason}){
    const uid=requireAdmin(auth),path='tenants/'+safeId(tenantId);
    if(!['active','suspended'].includes(status))fail('invalid-status');
    if(typeof reason!=='string'||reason.trim().length<3||reason.trim().length>250)fail('invalid-reason');
    const auditId=randomUUID();
    return store.runTransaction(async tx=>{
      const existing=await tx.get(path);
      if(!existing)fail('not-found');
      if(!['active','suspended'].includes(existing.status))fail('invalid-current-status');
      if(existing.status===status)return {changed:false,status};
      const now=clock();
      tx.set(path,{...existing,status,statusUpdatedAt:now,statusUpdatedBy:uid,
        inviteGeneration:(Number.isSafeInteger(existing.inviteGeneration)&&existing.inviteGeneration>=0?existing.inviteGeneration:0)+1});
      tx.create(path+'/audit/'+auditId,{event:'company-status-changed',actorUid:uid,previousStatus:existing.status,status,reason:reason.trim(),at:now});
      return {changed:true,status};
    });
  }
  return Object.freeze({listCompanies,setCompanyStatus});
}
module.exports={makePlatformAdminService};

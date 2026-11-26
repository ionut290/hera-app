'use strict';
// Tenant-scoped workspace read service. No global collection fallback or client roles.
const {readModule}=require('./modules.cjs');
const ROLES=new Set(['owner','admin','manager','team_lead','worker','client']);
function fail(code){throw Object.assign(new Error(code),{code});}
function id(value){if(typeof value!=='string'||!/^[A-Za-z0-9_-]{1,128}$/.test(value)||/^__.*__$/.test(value))fail('invalid-id');return value;}
function page(payload){const limit=payload.limit??25,after=payload.after??null;if(!Number.isInteger(limit)||limit<1||limit>50)fail('invalid-limit');if(after!==null)id(after);return {limit,after};}
function createWorkspaceService(db){
  async function membership(tx,uid,tenantId){
    const root=db.collection('tenants').doc(id(tenantId));
    const [company,member]=await Promise.all([tx.get(root),tx.get(root.collection('members').doc(id(uid)))]);
    if(!company.exists||!member.exists||company.data().status!=='active'||member.data().status!=='active'||!ROLES.has(member.data().role))return null;
    return {root,company:{id:root.id,name:String(company.data().name||root.id),role:member.data().role}};
  }
  async function listMyCompanies(caller,payload){
    const {limit,after}=page(payload),uid=id(caller.uid);
    return db.runTransaction(async tx=>{
      let query=db.collection('users').doc(uid).collection('memberships').orderBy('__name__').limit(limit+1);
      if(after)query=query.startAfter(after);
      const snap=await tx.get(query),rows=snap.docs.slice(0,limit);
      const checks=await Promise.all(rows.filter(d=>d.data().status==='active').map(d=>membership(tx,uid,d.id)));
      return {companies:checks.filter(Boolean).map(x=>x.company),nextCursor:snap.size>limit?rows.at(-1).id:null};
    });
  }
  async function openCompany(caller,payload){
    return db.runTransaction(async tx=>{
      const access=await membership(tx,caller.uid,payload.tenantId);
      if(!access)fail('forbidden');
      return {company:access.company,capabilities:{readCommesse:['owner','admin'].includes(access.company.role),readModules:['owner','admin'].includes(access.company.role)}};
    });
  }
  async function listCompanyCommesse(caller,payload){
    const {limit,after}=page(payload);
    if(!['ordinary','snow'].includes(payload.kind))fail('invalid-kind');
    return db.runTransaction(async tx=>{
      const access=await membership(tx,caller.uid,payload.tenantId);
      if(!access||!['owner','admin'].includes(access.company.role))fail('forbidden');
      const collection=payload.kind==='snow'?'neve_commesse':'commesse';
      let query=access.root.collection(collection).orderBy('__name__').limit(limit+1);
      if(after)query=query.startAfter(after);
      const snap=await tx.get(query),rows=snap.docs.slice(0,limit);
      return {company:access.company,kind:payload.kind,commesse:rows.map(d=>({id:d.id,name:String(d.data().nome??d.data().name??d.data().NOME??d.id)})),nextCursor:snap.size>limit?rows.at(-1).id:null};
    });
  }
  async function listCompanyModule(caller,payload){
    return db.runTransaction(async tx=>{
      const access=await membership(tx,caller.uid,payload.tenantId);
      if(!access||!['owner','admin'].includes(access.company.role))fail('forbidden');
      return readModule({tx,access,caller,payload,page,id});
    });
  }
  return Object.freeze({listMyCompanies,openCompany,listCompanyCommesse,listCompanyModule});
}
module.exports={createWorkspaceService};

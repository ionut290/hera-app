'use strict';
// Shared server-side tenant dispatcher. Never imported by browser or native bundle.
const {makeTenantService}=require('./tenant-service.cjs');
const {makePlatformAdminService}=require('./platform-admin.cjs');
const {createWorkspaceService}=require('./workspace.cjs');
const {service:starterService}=require('./starter-plan.cjs');
function fail(code){throw Object.assign(new Error(code),{code});}
function createTenantDispatcher({auth,db,assertBoundary}){
  if(!auth||!db||typeof assertBoundary!=='function')fail('invalid-adapter');
  assertBoundary();
  const store={
    runTransaction:fn=>db.runTransaction(t=>fn({
      get:async p=>{const s=await t.get(db.doc(p));return s.exists?s.data():null;},
      create:(p,d)=>t.create(db.doc(p),d),
      set:(p,d)=>t.set(db.doc(p),d)
    })),
    listTenants:async({limit,after})=>{let q=db.collection('tenants').orderBy('__name__').limit(limit+1);if(after)q=q.startAfter(after);return(await q.get()).docs.map(d=>({id:d.id,data:d.data()}));}
  };
  const tenant=makeTenantService(store),platform=makePlatformAdminService(store),workspace=createWorkspaceService(db),starter=starterService(db);
  async function dispatch(token,action,payload={}){
    assertBoundary();
    if(typeof token!=='string'||!token||token.length>16000)fail('unauthenticated');
    let decoded,user;
    try{
      decoded=await auth.verifyIdToken(token,true);
      user=await auth.getUser(decoded.uid);
    }catch{fail('unauthenticated');}
    if(user.disabled||!user.emailVerified||decoded.email_verified!==true||!user.email||decoded.email!==user.email||decoded.apple_review===true||user.customClaims?.apple_review===true)fail('unauthenticated');
    // Never trust role/UID/email fields from request payload or stale admin claims.
    const caller={uid:user.uid,email:user.email,email_verified:true,
      platform_admin:decoded.platform_admin===true&&user.customClaims?.platform_admin===true};
    if(!payload||typeof payload!=='object'||Array.isArray(payload))fail('invalid-payload');
    switch(action){
      case 'registerCompany':return starter.register(caller,payload);
      case 'getCompanyPlan':return starter.get(caller,payload);
      case 'requestPlanUpgrade':return starter.request(caller,payload);
      case 'listPendingPlanUpgrades':return starter.listPending(caller,payload);
      case 'getPlanForAdmin':return starter.get(caller,payload,true);
      case 'reviewPlanUpgrade':return starter.review(caller,payload);
      case 'listMyCompanies':return workspace.listMyCompanies(caller,payload);
      case 'openCompany':return workspace.openCompany(caller,payload);
      case 'listCompanyCommesse':return workspace.listCompanyCommesse(caller,payload);
      case 'listCompanyMembers':return workspace.listCompanyMembers(caller,payload);
      case 'listCompanyModule':return workspace.listCompanyModule(caller,payload);
      case 'getCompanyRecord':return workspace.getCompanyRecord(caller,payload);
      case 'saveCompanyRecord':return workspace.saveCompanyRecord(caller,payload);
      case 'createCompanyOperation':return workspace.createCompanyOperation(caller,payload);
      case 'createCompany':{
        if(!caller.platform_admin)fail('forbidden');
        if(typeof payload.ownerEmail!=='string')fail('invalid-email');
        let owner;
        try{owner=await auth.getUserByEmail(payload.ownerEmail.trim());}catch{fail('invalid-owner');}
        if(owner.disabled||!owner.emailVerified||owner.customClaims?.apple_review===true)fail('invalid-owner');
        return tenant.createCompany(caller,{name:payload.name,ownerUid:owner.uid,ownerEmail:owner.email,approvedLimits:payload.approvedLimits});
      }
      case 'invite':return tenant.invite(caller,{tenantId:payload.tenantId,recipientEmail:payload.recipientEmail,role:payload.role});
      case 'accept':return tenant.accept(caller,{tenantId:payload.tenantId,token:payload.inviteToken});
      case 'listCompanies':return platform.listCompanies(caller,{limit:payload.limit,after:payload.after});
      case 'setCompanyStatus':return platform.setCompanyStatus(caller,{tenantId:payload.tenantId,status:payload.status,reason:payload.reason});
      default:fail('unknown-action');
    }
  }
  return Object.freeze({dispatch});
}
module.exports={createTenantDispatcher};

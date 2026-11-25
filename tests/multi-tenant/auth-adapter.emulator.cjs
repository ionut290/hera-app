'use strict';
// Local integration adapter. Never import in production; fail closed outside dedicated emulators.
const {initializeApp,deleteApp}=require('firebase-admin/app');
const {getAuth}=require('firebase-admin/auth');
const {getFirestore}=require('firebase-admin/firestore');
const {makeTenantService}=require('./tenant-service.prototype.cjs');
const {makePlatformAdminService}=require('./platform-admin.prototype.cjs');
const PROJECT='demo-varga-multi-tenant';
function fail(code){throw Object.assign(new Error(code),{code});}
function validateEmulators(){
  if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9199')fail('emulator-only');
}
function createLocalAdapter(){
  validateEmulators();
  const app=initializeApp({projectId:PROJECT},'tenant-adapter-'+require('node:crypto').randomUUID());
  const db=getFirestore(app),auth=getAuth(app);
  const store={
    runTransaction:fn=>db.runTransaction(t=>fn({
      get:async p=>{const s=await t.get(db.doc(p));return s.exists?s.data():null;},
      create:(p,d)=>t.create(db.doc(p),d),
      set:(p,d)=>t.set(db.doc(p),d)
    })),
    listTenants:async({limit,after})=>{let q=db.collection('tenants').orderBy('__name__').limit(limit+1);if(after)q=q.startAfter(after);return(await q.get()).docs.map(d=>({id:d.id,data:d.data()}));}
  };
  const tenant=makeTenantService(store),platform=makePlatformAdminService(store);
  async function dispatch(token,action,payload={}){
    validateEmulators();
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
      case 'createCompany':{
        if(!caller.platform_admin)fail('forbidden');
        if(typeof payload.ownerEmail!=='string')fail('invalid-email');
        let owner;
        try{owner=await auth.getUserByEmail(payload.ownerEmail.trim());}catch{fail('invalid-owner');}
        if(owner.disabled||!owner.emailVerified||owner.customClaims?.apple_review===true)fail('invalid-owner');
        return tenant.createCompany(caller,{name:payload.name,ownerUid:owner.uid,ownerEmail:owner.email});
      }
      case 'invite':return tenant.invite(caller,{tenantId:payload.tenantId,recipientEmail:payload.recipientEmail,role:payload.role});
      case 'accept':return tenant.accept(caller,{tenantId:payload.tenantId,token:payload.inviteToken});
      case 'listCompanies':return platform.listCompanies(caller,{limit:payload.limit,after:payload.after});
      case 'setCompanyStatus':return platform.setCompanyStatus(caller,{tenantId:payload.tenantId,status:payload.status,reason:payload.reason});
      default:fail('unknown-action');
    }
  }
  return {dispatch,auth,db,close:async()=>{await db.terminate();await deleteApp(app);}};
}
module.exports={createLocalAdapter,PROJECT};

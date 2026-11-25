'use strict';
// Explicitly live STAGING only. No production deployment, service-account keys,
// persistent tokens, client-side admin grants, public endpoints or paid resources.
const {execFileSync}=require('node:child_process');
const {randomUUID,randomBytes}=require('node:crypto');
const assert=require('node:assert/strict');
const PROJECT='varga-cantieri-v2-staging';
const API_KEY='AIzaSyC3ujdb0XyHKk-B0eGt8IVr_U_DKEXxaLI'; // public Firebase Web configuration
const ROOT='projects/'+PROJECT+'/databases/(default)/documents';
const FS='https://firestore.googleapis.com/v1/'+ROOT;
const AUTH='https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key='+API_KEY;
const MARKER='stagingBootstrap/v2-demo-initial';
function fail(code){throw Object.assign(new Error(code),{code});}
function gcloud(args){return execFileSync('gcloud',args,{encoding:'utf8',stdio:['ignore','pipe','ignore'],timeout:30000}).trim();}
function guard(args){
  if(args.length!==3||args[0]!=='--execute'||args[1]!==PROJECT||!/^[^\s@]+@[^\s@]+$/.test(args[2]))fail('explicit-staging-and-account-required');
  for(const key of ['FIRESTORE_EMULATOR_HOST','FIREBASE_AUTH_EMULATOR_HOST','GOOGLE_APPLICATION_CREDENTIALS'])if(process.env[key])fail('unexpected-credential-or-emulator-override');
  for(const key of ['GOOGLE_CLOUD_PROJECT','GCLOUD_PROJECT','CLOUDSDK_CORE_PROJECT'])if(process.env[key]&&process.env[key]!==PROJECT)fail('wrong-environment-project');
  if(gcloud(['auth','list','--filter=status:ACTIVE','--format=value(account)'])!==args[2])fail('wrong-google-account');
  const p=JSON.parse(gcloud(['projects','describe',PROJECT,'--account='+args[2],'--format=json','--quiet']));
  if(p.projectId!==PROJECT||p.lifecycleState!=='ACTIVE')fail('wrong-project');
  const billing=JSON.parse(gcloud(['billing','projects','describe',PROJECT,'--account='+args[2],'--format=json','--quiet']));
  if(billing.billingEnabled!==false||billing.billingAccountName)fail('billing-not-confirmed-off');
  return args[2];
}
function fields(data){return Object.fromEntries(Object.entries(data).map(([k,v])=>[k,typeof v==='boolean'?{booleanValue:v}:typeof v==='number'?{integerValue:String(v)}:{stringValue:String(v)}]));}
async function request(url,token,method='GET',body){
  const headers={'content-type':'application/json'};if(token)headers.authorization='Bearer '+token;
  const response=await fetch(url,{method,headers,body:body===undefined?undefined:JSON.stringify(body),signal:AbortSignal.timeout(20000),redirect:'error'});
  const data=await response.json().catch(()=>({}));return {status:response.status,data};
}
function requireStatus(r,expected,label){if(r.status!==expected)fail(label+'-http-'+r.status);return r.data;}
function createWrite(path,data){return {update:{name:ROOT+'/'+path,fields:fields(data)},currentDocument:{exists:false}};}
async function run(args){
  const account=guard(args);
  const {initializeApp,deleteApp}=require('firebase-admin/app');
  const {getAuth}=require('firebase-admin/auth');
  const access=()=>gcloud(['auth','print-access-token','--account='+account,'--quiet']);
  const app=initializeApp({projectId:PROJECT,credential:{getAccessToken:async()=>({access_token:access(),expires_in:300})}},'staging-bootstrap');
  const auth=getAuth(app),created=[];let committed=false,cleanupFailed=false;
  try{
    const marker=await request(FS+'/'+MARKER,access());
    requireStatus(marker,404,'bootstrap-marker-already-exists-or-unavailable');
    // No overwrite or recovery of any pre-existing account. Every UID belongs to this run only.
    const runId=randomUUID(),now=Date.now(),prefix='v2-demo-'+runId;
    const people=[];
    for(const role of ['platform','owner-a','owner-b']){
      const uid=prefix+'-'+role,email=uid+'@example.invalid',password=randomBytes(32).toString('base64url')+'!aA9';
      const user=await auth.createUser({uid,email,password,emailVerified:true,displayName:'STAGING synthetic '+role});
      created.push(user.uid);
      if(role==='platform')await auth.setCustomUserClaims(uid,{platform_admin:true,staging_synthetic:true});
      else await auth.setCustomUserClaims(uid,{staging_synthetic:true});
      const signed=requireStatus(await request(AUTH,null,'POST',{email,password,returnSecureToken:true}),200,'synthetic-sign-in');
      const decoded=await auth.verifyIdToken(signed.idToken,true);
      assert.equal(decoded.aud,PROJECT);assert.equal(decoded.uid,uid);assert.equal(decoded.email_verified,true);assert.equal(decoded.platform_admin===true,role==='platform');
      people.push({uid,email,token:signed.idToken});
    }
    const [admin,a,b]=people;
    const companies=[{id:prefix+'-a',name:'DEMO Verde',owner:a},{id:prefix+'-b',name:'DEMO Edilizia',owner:b}];
    const writes=[createWrite(MARKER,{runId,createdAt:now,synthetic:true,tenantA:companies[0].id,tenantB:companies[1].id})];
    for(const c of companies){
      const root='tenants/'+c.id;
      writes.push(createWrite(root,{name:c.name,status:'active',inviteGeneration:0,createdAt:now,createdBy:admin.uid,synthetic:true}),
        createWrite(root+'/members/'+c.owner.uid,{role:'owner',status:'active',email:c.owner.email,createdAt:now}),
        createWrite('users/'+c.owner.uid+'/memberships/'+c.id,{role:'owner',status:'active'}),
        createWrite(root+'/audit/'+runId,{event:'company-created',actorUid:admin.uid,subjectUid:c.owner.uid,at:now,synthetic:true}));
    }
    requireStatus(await request(FS+':commit',access(),'POST',{writes}),200,'atomic-synthetic-create');committed=true;
    console.log('Created two synthetic companies; three synthetic Auth users.');
    let passed=0;
    async function check(url,token,status,label,method='GET',body){requireStatus(await request(url,token,method,body),status,label);passed++;}
    const own=companies[0].id,other=companies[1].id;
    await check(FS+'/tenants/'+own,a.token,200,'owner-a-own');
    await check(FS+'/tenants/'+other,b.token,200,'owner-b-own');
    await check(FS+'/tenants/'+other,a.token,403,'owner-a-cross');
    await check(FS+'/tenants/'+own,b.token,403,'owner-b-cross');
    await check(FS+'/tenants/'+own+'/members/'+a.uid,a.token,200,'owner-a-membership');
    await check(FS+'/users/'+a.uid+'/memberships',a.token,200,'own-index');
    await check(FS+'/users/'+b.uid+'/memberships',a.token,403,'cross-index');
    await check(FS+'/tenants/'+own,admin.token,403,'platform-not-member');
    await check(FS+'/tenants/'+own,null,403,'anonymous-denied');
    await check(FS+'/tenants/'+own,a.token,403,'client-write-denied','PATCH',{fields:{name:{stringValue:'Synthetic forbidden write'}}});
    await check(FS+'/tenants/'+own+'/audit/'+runId,a.token,403,'audit-denied');
    console.log('LIVE_FIRESTORE_CHECKS_PASSED='+passed);
    console.log('Synthetic tenant IDs: '+own+' '+other);
  }finally{
    // Leave synthetic docs for inspection. Revoke claims and disable only users created here.
    // No passwords or ID tokens are printed, persisted or committed.
    for(const uid of created){
      try{await auth.updateUser(uid,{disabled:true});await auth.setCustomUserClaims(uid,{});await auth.revokeRefreshTokens(uid);}
      catch{cleanupFailed=true;console.error('SYNTHETIC_ACCOUNT_CLEANUP_REQUIRED='+uid);}
    }
    await deleteApp(app);
    console.log('SYNTHETIC_DOCS_COMMITTED='+committed);
    console.log('SYNTHETIC_ACCOUNTS_DISABLED='+(!cleanupFailed));
    if(cleanupFailed)fail('synthetic-account-cleanup-incomplete');
  }
}
if(require.main===module)run(process.argv.slice(2)).catch(e=>{console.error('STAGING_BOOTSTRAP_FAILED:',e.code||e.name);process.exitCode=1;});
module.exports={fields,createWrite,ROOT};

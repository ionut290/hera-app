'use strict';
// All newly provisioned companies have an explicit quota policy. Missing policy
// cannot authorize a quota-consuming mutation. Existing production data untouched.
const {randomUUID}=require('node:crypto');
const DEFAULT=Object.freeze({jobs:1,plants:10,admins:1,operators:3});
const keys=Object.keys(DEFAULT),usageKeys=[...keys,'personnel'];
const fail=code=>{throw Object.assign(new Error(code),{code});};
function starter(approvedLimits){return {version:1,limits:approvedLimits==null?{...DEFAULT}:limitInput(approvedLimits,DEFAULT,true),usage:{jobs:0,plants:0,admins:1,operators:0,personnel:0},pending:null};}
function policy(data){
  if(data?.starterPlan==null)return null;
  const p=data.starterPlan;
  if(p.version!==1||!p.limits||!p.usage||keys.some(k=>!Number.isSafeInteger(p.limits[k])||p.limits[k]<DEFAULT[k])||usageKeys.some(k=>!Number.isSafeInteger(p.usage[k])||p.usage[k]<0))fail('invalid-plan');
  return p;
}
function increment(data,key,checkOnly=false){
  const p=policy(data);if(!p)fail('invalid-plan-unavailable');
  if(!usageKeys.includes(key))fail('invalid-plan');
  const limit=p.limits[key==='personnel'?'operators':key];
  if(p.usage[key]>=limit)fail('invalid-quota-'+key);
  return checkOnly?null:{...p,usage:{...p.usage,[key]:p.usage[key]+1}};
}
function limitInput(value,old,allowEqual=false){
  if(!value||typeof value!=='object'||Array.isArray(value)||Object.keys(value).length!==keys.length||keys.some(k=>!Number.isSafeInteger(value[k])||value[k]<old[k]||value[k]>100000))fail('invalid-limits');
  if(!allowEqual&&!keys.some(k=>value[k]>old[k]))fail('invalid-limits');
  return Object.fromEntries(keys.map(k=>[k,value[k]]));
}
function safeId(value){if(typeof value!=='string'||!/^[A-Za-z0-9_-]{1,128}$/.test(value)||/^__.*__$/.test(value))fail('invalid-id');return value;}
function reason(value){if(typeof value!=='string'||value.trim().length<3||value.trim().length>250)fail('invalid-reason');return value.trim();}
function service(db){
  async function register(caller,payload){
    if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9199')fail('forbidden');
    if(typeof payload.name!=='string'||!payload.name.trim()||payload.name.trim().length>100)fail('invalid-name');
    if(Object.keys(payload).some(k=>k!=='name'))fail('invalid-payload');
    const uid=safeId(caller.uid),claim=db.collection('selfServiceCompanies').doc(uid),id=randomUUID(),root=db.collection('tenants').doc(id),now=new Date();
    return db.runTransaction(async tx=>{
      const previous=await tx.get(claim);
      if(previous.exists)return {tenantId:previous.data().tenantId,alreadyRegistered:true};
      const membership={role:'owner',status:'active',email:caller.email,createdAt:now};
      tx.create(root,{name:payload.name.trim(),status:'active',inviteGeneration:0,createdAt:now,createdBy:uid,starterPlan:starter(),starterUpgradePending:false});
      tx.create(root.collection('members').doc(uid),membership);
      tx.create(db.collection('users').doc(uid).collection('memberships').doc(id),{role:'owner',status:'active'});
      tx.create(claim,{tenantId:id,createdAt:now});
      tx.create(root.collection('audit').doc(),{event:'starter-registered',actorUid:uid,at:now});
      return {tenantId:id,alreadyRegistered:false};
    });
  }
  async function access(tx,caller,payload,superOnly=false){
    const root=db.collection('tenants').doc(safeId(payload.tenantId)),snap=await tx.get(root);
    if(superOnly){if(!caller.platform_admin)fail('forbidden');}
    else{
      const m=await tx.get(root.collection('members').doc(safeId(caller.uid)));
      if(!m.exists||m.data().status!=='active'||!['owner','admin'].includes(m.data().role)||snap.data()?.status!=='active')fail('forbidden');
    }
    if(!snap.exists)fail('not-found');
    const plan=policy(snap.data());
    if(!plan)fail('invalid-plan-unavailable');
    return {root,plan};
  }
  async function get(caller,payload,superOnly=false){return db.runTransaction(async tx=>{const {plan}=await access(tx,caller,payload,superOnly);return {plan};});}
  async function request(caller,payload){
    return db.runTransaction(async tx=>{
      const {root,plan}=await access(tx,caller,payload);
      if(plan.pending)fail('invalid-pending-request');
      const limits=limitInput(payload.limits,plan.limits),why=reason(payload.reason);
      const pending={id:randomUUID(),limits,reason:why,requestedBy:caller.uid,requestedAt:new Date()};
      tx.update(root,{starterPlan:{...plan,pending},starterUpgradePending:true});
      tx.create(root.collection('audit').doc(),{event:'starter-upgrade-requested',requestId:pending.id,actorUid:caller.uid,limits,at:new Date()});
      return {requestId:pending.id};
    });
  }
  async function review(caller,payload){
    if(!caller.platform_admin)fail('forbidden');
    return db.runTransaction(async tx=>{
      const {root,plan}=await access(tx,caller,payload,true);
      if(!plan.pending||plan.pending.id!==safeId(payload.requestId))fail('invalid-request');
      if(!['approve','reject'].includes(payload.decision))fail('invalid-decision');
      const why=reason(payload.reason);
      const limits=payload.decision==='approve'?limitInput(payload.limits??plan.pending.limits,plan.limits):plan.limits;
      tx.update(root,{starterPlan:{...plan,limits,pending:null,lastDecision:{requestId:plan.pending.id,decision:payload.decision,reason:why,at:new Date()}},starterUpgradePending:false});
      tx.create(root.collection('audit').doc(),{event:'starter-upgrade-'+payload.decision,actorUid:caller.uid,requestId:plan.pending.id,previousLimits:plan.limits,limits,reason:why,at:new Date()});
      return {limits,decision:payload.decision};
    });
  }
  async function listPending(caller,payload={}){
    if(!caller.platform_admin)fail('forbidden');
    const limit=payload.limit??25,after=payload.after??null;
    if(!Number.isSafeInteger(limit)||limit<1||limit>50)fail('invalid-limit');
    if(after!==null)safeId(after);
    let query=db.collection('tenants').where('starterUpgradePending','==',true).orderBy('__name__').limit(limit+1);
    if(after)query=query.startAfter(after);
    const snap=await query.get(),rows=snap.docs.slice(0,limit);
    const requests=rows.flatMap(d=>{const p=policy(d.data());return p?.pending?[{tenantId:d.id,name:String(d.data().name||d.id),limits:p.limits,requested:p.pending.limits,requestId:p.pending.id,reason:p.pending.reason}]:[];});
    return {requests,nextCursor:snap.size>limit?rows.at(-1).id:null};
  }
  return {register,get,request,review,listPending};
}
module.exports={starter,policy,increment,service};

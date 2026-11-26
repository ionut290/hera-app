'use strict';
// Self-service hours pilot, local emulator only. Reuses the existing pending
// hours operation and its person/job/date uniqueness lock. No approval or FATTO.
const {createHash}=require('node:crypto');
const {createOperation}=require('./operations.cjs');
const fail=code=>{throw Object.assign(new Error(code),{code});};
const key=parts=>createHash('sha256').update(JSON.stringify(parts)).digest('hex');
function local(){if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9199')fail('forbidden');}
function kindOf(k){if(!['ordinary','snow'].includes(k))fail('invalid-kind');return k;}
function personRef(root,k,p){return root.collection(k==='snow'?'neve_personale':'personale').doc(p);}
function binding(root,uid,k){return root.collection('members').doc(uid).collection('personnelBindings').doc(k);}
function reverse(root,k,p){return root.collection('tenantLabPersonnelBindings').doc(key([k,p]));}
function allowFields(payload,fields){if(Object.keys(payload).some(k=>!fields.includes(k)))fail('invalid-payload');}
async function bindWorker({tx,access,caller,payload,id}){
  local();allowFields(payload,['tenantId','kind','memberUid','personnelId']);
  const root=access.root,k=kindOf(payload.kind),uid=id(payload.memberUid),p=id(payload.personnelId);
  const a=binding(root,uid,k),r=reverse(root,k,p);
  const [m,person,old,back]=await Promise.all([tx.get(root.collection('members').doc(uid)),tx.get(personRef(root,k,p)),tx.get(a),tx.get(r)]);
  if(!m.exists||m.data().status!=='active'||m.data().role!=='worker')fail('invalid-member');
  if(!person.exists||person.data().tenantLabManaged!==true)fail('invalid-reference');
  // Do not reassign existing historical links or silently move a person to another user.
  if(old.exists||back.exists){
    if(old.exists&&back.exists&&old.data().personnelId===p&&back.data().memberUid===uid)return {changed:false,kind:k,memberUid:uid,personnelId:p};
    fail('invalid-binding-conflict');
  }
  const now=new Date();
  tx.create(a,{kind:k,personnelId:p,createdBy:caller.uid,createdAt:now});
  tx.create(r,{kind:k,memberUid:uid,personnelId:p,createdAt:now});
  tx.create(root.collection('audit').doc(),{event:'lab-worker-personnel-linked',kind:k,memberUid:uid,personnelId:p,actorUid:caller.uid,at:now});
  return {changed:true,kind:k,memberUid:uid,personnelId:p};
}
async function submitHours({tx,access,caller,payload,id}){
  local();if(access.company.role!=='worker')fail('forbidden');
  allowFields(payload,['tenantId','kind','commessaId','plantId','requestId','date','hours']);
  const root=access.root,uid=id(caller.uid),k=kindOf(payload.kind),j=id(payload.commessaId),p=id(payload.plantId);
  const job=root.collection(k==='snow'?'neve_commesse':'commesse').doc(j);
  const [g,b,plant]=await Promise.all([tx.get(root.collection('members').doc(uid).collection('plantGrants').doc(key([k,j,p]))),tx.get(binding(root,uid,k)),tx.get(job.collection('impianti').doc(p))]);
  if(!g.exists||g.data().enabled!==true||g.data().kind!==k||g.data().jobId!==j||g.data().plantId!==p||!plant.exists||plant.data().tenantLabPlant!==true)fail('forbidden');
  if(!b.exists||b.data().kind!==k)fail('invalid-personnel-binding');
  const person=id(b.data().personnelId),back=await tx.get(reverse(root,k,person));
  if(!back.exists||back.data().memberUid!==uid||back.data().personnelId!==person||back.data().kind!==k)fail('invalid-personnel-binding');
  // All reads above precede the shared operation's reads and writes.
  // Caller remains the verified worker; no impersonation of an administrator.
  return createOperation({tx,access,caller,id,payload:{operation:'hoursRequest',kind:k,commessaId:j,requestId:payload.requestId,data:{date:payload.date,hours:payload.hours,personnelId:person}}});
}
async function listMyHours({tx,access,caller,payload,page,id}){
  local();if(access.company.role!=='worker')fail('forbidden');
  allowFields(payload,['tenantId','kind','limit','after']);
  const k=kindOf(payload.kind),{limit,after}=page(payload),uid=id(caller.uid);
  let q=access.root.collection(k==='snow'?'neve_ore_richieste':'oreApprovalRequests').where('createdByUid','==',uid).orderBy('__name__').limit(limit+1);
  if(after)q=q.startAfter(after);
  const snap=await tx.get(q),docs=snap.docs.slice(0,limit);
  const rows=docs.flatMap(d=>{const x=d.data();if(x.tenantLabHours!==true||!Array.isArray(x.entries)||x.entries.length!==1||!Array.isArray(x.entries[0].rows)||x.entries[0].rows.length!==1)return [];
    const row=x.entries[0].rows[0];return [{id:d.id,date:String(x.date||''),commessa:String(x.commessaName||'').slice(0,150),hours:typeof row.ore==='number'?row.ore:null,status:String(x.status||'')}];});
  return {company:access.company,kind:k,rows,nextCursor:snap.size>limit?docs.at(-1).id:null};
}
module.exports={bindWorker,submitHours,listMyHours};

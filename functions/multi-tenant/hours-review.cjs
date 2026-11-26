'use strict';
// Local pilot of original pending_level1 -> pending_admin -> approved/rejected.
// Atomically preserve request, receipt/lock and report. Never send live notifications.
const {createHash}=require('node:crypto');
const fail=code=>{throw Object.assign(new Error(code),{code});};
const key=x=>createHash('sha256').update(JSON.stringify(x)).digest('hex');
async function reviewHours({tx,access,caller,payload,id}){
  if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9199')fail('forbidden');
  if(Object.keys(payload).some(k=>!['tenantId','kind','requestId','expectedStatus','decision','reason'].includes(k)))fail('invalid-payload');
  const {kind,decision,expectedStatus}=payload,requestId=id(payload.requestId);
  if(!['ordinary','snow'].includes(kind)||!['approve','reject'].includes(decision)||!['pending_level1','pending_admin'].includes(expectedStatus))fail('invalid-payload');
  const reason=payload.reason??'';
  if(typeof reason!=='string'||reason.length>250)fail('invalid-reason');
  const root=access.root,snow=kind==='snow',ref=root.collection(snow?'neve_ore_richieste':'oreApprovalRequests').doc(requestId),snap=await tx.get(ref);
  if(!snap.exists||snap.data().tenantLabHours!==true||snap.data().tenantLabManaged!==true)fail('not-found');
  const x=snap.data(),admin=['owner','admin'].includes(access.company.role);
  if(!admin&&!(access.company.role==='worker'&&expectedStatus==='pending_level1'&&x.createdByUid===caller.uid))fail('forbidden');
  if(x.status!==expectedStatus)fail('invalid-hours-status');
  if(expectedStatus==='pending_admin'&&!admin)fail('forbidden');
  if(!Array.isArray(x.entries)||x.entries.length!==1||!Array.isArray(x.entries[0].rows)||x.entries[0].rows.length!==1)fail('invalid-hours-review');
  const entry=x.entries[0],row=entry.rows[0],job=id(entry.commessaId),person=id(row.operatoreId);
  if(typeof x.date!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(x.date)||typeof row.ore!=='number'||row.ore<=0||row.ore>24||Math.round(row.ore*4)!==row.ore*4)fail('invalid-hours-review');
  const lock=root.collection('tenantLabHoursLocks').doc(key([x.date,job,person]));
  const report=root.collection(snow?'neve_ore':'oreReports').doc(requestId);
  const [l,p,j,existing]=await Promise.all([tx.get(lock),tx.get(root.collection(snow?'neve_personale':'personale').doc(person)),tx.get(root.collection(snow?'neve_commesse':'commesse').doc(job)),tx.get(report)]);
  if(!l.exists||l.data().requestId!==requestId||l.data().status!==expectedStatus||l.data().personnelId!==person||l.data().commessaId!==job||l.data().date!==x.date||!p.exists||p.data().tenantLabManaged!==true||!j.exists||j.data().tenantLabManaged!==true||existing.exists)fail('invalid-hours-review');
  const now=new Date(),actor=caller.uid;
  const status=decision==='reject'?'rejected':expectedStatus==='pending_level1'?'pending_admin':'approved';
  let patch={status,updatedAt:now};
  if(status==='pending_admin')patch={...patch,level1ApprovedBy:caller.email,level1ApprovedByUid:actor,level1ApprovedAt:now};
  if(status==='rejected')patch={...patch,rejectedBy:caller.email,rejectedByUid:actor,rejectedAt:now,rejectionReason:reason.trim()};
  if(status==='approved'){
    if(!x.level1ApprovedAt||!x.level1ApprovedBy)fail('invalid-hours-review');
    patch={...patch,level2ApprovedBy:caller.email,level2ApprovedByUid:actor,level2ApprovedAt:now,finalizedReportId:report.id};
    tx.create(report,{date:x.date,entries:x.entries,commessaName:x.commessaName,createdByUid:x.createdByUid,createdByEmail:x.createdByEmail,createdAt:now,sourceApprovalRequestId:requestId,tenantLabHours:true,tenantLabManaged:true});
  }
  tx.update(ref,patch);tx.update(lock,{status,updatedAt:now,...(status==='approved'?{reportId:report.id}:{})});
  tx.create(root.collection('audit').doc(),{event:'lab-hours-'+status,requestId,kind,actorUid:actor,at:now,previousStatus:expectedStatus});
  return {company:access.company,requestId,status,reportId:status==='approved'?report.id:null};
}
module.exports={reviewHours};

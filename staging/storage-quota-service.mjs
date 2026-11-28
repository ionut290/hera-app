// Candidate backend only. No live deployment or public download URLs.
// All writes use an Admin SDK identity; client Storage rules must deny direct access.
import {createHash} from 'node:crypto';
export const MAX_FILE_BYTES=15*1024*1024;
const PROJECT='varga-cantieri-v2-staging',DAY=86400000;
const idOk=s=>typeof s==='string'&&/^[a-zA-Z0-9_-]{16,80}$/.test(s);
const tenantOk=s=>typeof s==='string'&&/^demo-[a-z0-9-]{3,59}$/.test(s);
const types=new Set(['application/pdf','image/jpeg','image/png','image/webp']);
const fail=code=>{throw Object.assign(new Error(code),{code});};
function writable(t,now){
 if(t.plan==='trial')return t.trialDays===30&&Number.isFinite(t.trialStartedAt?.toMillis?.())&&now<t.trialStartedAt.toMillis()+30*DAY;
 if(t.plan==='full'&&t.billing)return Number.isSafeInteger(t.billing.periods)&&t.billing.periods>0&&t.billing.periods<=120&&Number.isFinite(t.billing.periodStart?.toMillis?.())&&now<t.billing.periodStart.toMillis()+t.billing.periods*30*DAY;
 return t.plan==='full'&&!t.billing;
}
function validateFile({name,type,base64}){
 if(typeof name!=='string'||!name.trim()||name.length>120||/[\/\\\x00-\x1f\x7f]/.test(name)||!types.has(type))fail('INVALID_FILE');
 if(typeof base64!=='string'||base64.length>Math.ceil(MAX_FILE_BYTES/3)*4||base64.length%4!==0||!/^[A-Za-z0-9+/]+={0,2}$/.test(base64))fail('INVALID_FILE');
 const bytes=Buffer.from(base64,'base64');
 if(!bytes.length||bytes.length>MAX_FILE_BYTES||bytes.toString('base64')!==base64)fail('INVALID_FILE');
 const signature=type==='application/pdf'?bytes.subarray(0,5).toString()==='%PDF-':type==='image/jpeg'?bytes.subarray(0,3).equals(Buffer.from([255,216,255])):type==='image/png'?bytes.subarray(0,8).equals(Buffer.from([137,80,78,71,13,10,26,10])):bytes.subarray(0,4).toString()==='RIFF'&&bytes.subarray(8,12).toString()==='WEBP';
 if(!signature)fail('INVALID_FILE_TYPE');
 return {bytes,sha256:createHash('sha256').update(bytes).digest('hex')};
}
export function createStorageQuotaService({db,auth,bucket,projectId,serverTimestamp}){
 if(projectId!==PROJECT||db?.app?.options?.projectId!==PROJECT||auth?.app?.options?.projectId!==PROJECT||!bucket?.name?.startsWith(PROJECT+'.')||typeof serverTimestamp!=='function')fail('STAGING_ONLY');
 async function identity(token){
  if(typeof token!=='string'||!token)fail('UNAUTHENTICATED');
  const claims=await auth.verifyIdToken(token,true);
  if(!claims.uid||claims.email_verified!==true||claims.apple_review===true)fail('UNAUTHENTICATED');
  return claims.uid;
 }
 function refs(tenantId,uid,fileId){
  if(!tenantOk(tenantId)||!idOk(fileId))fail('INVALID_PATH');
  const tenant=db.doc('tenants/'+tenantId);
  return {tenant,member:tenant.collection('members').doc(uid),mirror:db.doc('users/'+uid+'/memberships/'+tenantId),meta:tenant.collection('storedFiles').doc(fileId),path:'tenants/'+tenantId+'/private-files/'+fileId};
 }
 async function membership(tx,r,tenantId){
  const [ts,ms,is]=await Promise.all([tx.get(r.tenant),tx.get(r.member),tx.get(r.mirror)]);
  const t=ts.data(),m=ms.data(),i=is.data();
  if(!t||t.demo!==true||t.status!=='active'||!m||m.status!=='active'||!i||i.status!=='active'||i.tenantId!==tenantId||m.role!==i.role||!['admin','supervisor','employee'].includes(m.role))fail('FORBIDDEN');
  return {t,m};
 }
 function view(d,id){return {id,name:d.name,type:d.type,size:d.size,status:d.status};}
 async function upload({token,tenantId,fileId,name,type,base64}){
  const uid=await identity(token),r=refs(tenantId,uid,fileId),{bytes,sha256}=validateFile({name,type,base64});
  const result=await db.runTransaction(async tx=>{
   const {t}=await membership(tx,r,tenantId);
   if(!writable(t,Date.now()))fail('PLAN_EXPIRED');
   const existing=(await tx.get(r.meta)).data();
   if(existing){
    if(existing.ownerUid!==uid||existing.sha256!==sha256||existing.name!==name||existing.type!==type||existing.size!==bytes.length)fail('FILE_ID_CONFLICT');
    if(existing.status!=='ready')fail('FILE_PENDING_RECONCILIATION');
    return {existing:true,view:view(existing,fileId)};
   }
   const used=t.usage?.storageBytes,limit=t.limits?.storageBytes;
   if(!Number.isSafeInteger(used)||used<0||!Number.isSafeInteger(limit)||limit<0||used>limit)fail('QUOTA_UNVERIFIED');
   if(bytes.length>limit-used)fail('STORAGE_LIMIT');
   const data={ownerUid:uid,name,type,size:bytes.length,sha256,status:'reserved',createdAt:serverTimestamp()};
   tx.update(r.tenant,{'usage.storageBytes':used+bytes.length});
   tx.create(r.meta,data);
   return {existing:false,view:view(data,fileId)};
  });
  if(result.existing)return result.view;
  // Reservation stays charged on any ambiguous error. Never release quota based
  // solely on a failed upload response: a blob may already exist.
  try{
   const file=bucket.file(r.path);
   await file.save(bytes,{resumable:false,validation:'crc32c',preconditionOpts:{ifGenerationMatch:0},metadata:{contentType:type,cacheControl:'private, no-store',metadata:{sha256}}});
   const [meta]=await file.getMetadata();
   if(Number(meta.size)!==bytes.length||meta.metadata?.sha256!==sha256)fail('BLOB_MISMATCH');
   await db.runTransaction(async tx=>{
    const d=(await tx.get(r.meta)).data();
    if(d?.status!=='reserved'||d.ownerUid!==uid||d.sha256!==sha256)fail('RESERVATION_CHANGED');
    tx.update(r.meta,{status:'ready',generation:String(meta.generation),readyAt:serverTimestamp()});
   });
   return {...result.view,status:'ready'};
  }catch{fail('FILE_PENDING_RECONCILIATION');}
 }
 async function download({token,tenantId,fileId}){
  const uid=await identity(token),r=refs(tenantId,uid,fileId);
  const d=await db.runTransaction(async tx=>{
   const {m}=await membership(tx,r,tenantId),d=(await tx.get(r.meta)).data();
   if(!d||d.status!=='ready'||(d.ownerUid!==uid&&m.role!=='admin'))fail('FORBIDDEN');
   return d;
  });
  const [bytes]=await bucket.file(r.path,{generation:d.generation}).download();
  if(bytes.length!==d.size||createHash('sha256').update(bytes).digest('hex')!==d.sha256)fail('BLOB_MISMATCH');
  return {...view(d,fileId),base64:bytes.toString('base64')};
 }
 return Object.freeze({upload,download});
}

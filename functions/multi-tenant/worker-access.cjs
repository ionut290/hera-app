'use strict';
// Local pilot: explicit per-plant grants, no legacy team/name inference or protected actions.
const {createHash}=require('node:crypto');
const fail=code=>{throw Object.assign(new Error(code),{code});};
const key=(kind,job,plant)=>createHash('sha256').update(JSON.stringify([kind,job,plant])).digest('hex');
const str=v=>typeof v==='string'?v.slice(0,250):'';
function path(root,kind,jobId,plantId){return root.collection(kind==='snow'?'neve_commesse':'commesse').doc(jobId).collection('impianti').doc(plantId);}
function checkLocal(){if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9199')fail('forbidden');}
async function setPlantGrant({tx,access,caller,payload,id}){
  checkLocal();
  const {kind,enabled}=payload,jobId=id(payload.commessaId),plantId=id(payload.plantId),uid=id(payload.memberUid);
  if(!['ordinary','snow'].includes(kind)||typeof enabled!=='boolean')fail('invalid-payload');
  const root=access.root,m=root.collection('members').doc(uid),job=root.collection(kind==='snow'?'neve_commesse':'commesse').doc(jobId);
  const grant=m.collection('plantGrants').doc(key(kind,jobId,plantId));
  const [member,j,p,old]=await Promise.all([tx.get(m),tx.get(job),tx.get(path(root,kind,jobId,plantId)),tx.get(grant)]);
  if(!member.exists||member.data().role!=='worker'||member.data().status!=='active')fail('invalid-member');
  if(!j.exists||j.data().tenantLabManaged!==true||!p.exists||p.data().tenantLabPlant!==true)fail('invalid-reference');
  if((old.exists&&old.data().enabled===enabled)||(!old.exists&&!enabled))return {changed:false,enabled};
  const now=new Date();
  tx.set(grant,{kind,jobId,plantId,enabled,updatedBy:caller.uid,updatedAt:now});
  tx.create(root.collection('audit').doc(),{event:enabled?'lab-plant-granted':'lab-plant-grant-revoked',memberUid:uid,kind,jobId,plantId,actorUid:caller.uid,at:now});
  return {changed:true,enabled};
}
async function listMyPlants({tx,access,caller,payload,page,id}){
  checkLocal();if(access.company.role!=='worker')fail('forbidden');
  const {limit,after}=page(payload),root=access.root;
  let query=root.collection('members').doc(id(caller.uid)).collection('plantGrants').where('enabled','==',true).orderBy('__name__').limit(limit+1);
  if(after)query=query.startAfter(after);
  const snap=await tx.get(query),grants=snap.docs.slice(0,limit);
  const rows=await Promise.all(grants.map(async grant=>{
    const g=grant.data();if(!['ordinary','snow'].includes(g.kind))return null;
    const jobId=id(g.jobId),plantId=id(g.plantId);
    if(grant.id!==key(g.kind,jobId,plantId))return null;
    const [job,plant]=await Promise.all([tx.get(root.collection(g.kind==='snow'?'neve_commesse':'commesse').doc(jobId)),tx.get(path(root,g.kind,jobId,plantId))]);
    if(!job.exists||job.data().tenantLabManaged!==true||!plant.exists||plant.data().tenantLabPlant!==true)return null;
    const p=plant.data();
    return {id:plantId,kind:g.kind,commessaId:jobId,commessa:str(job.data().nome),name:str(p.denominazione),comune:str(p.comune),indirizzo:str(p.indirizzo),status:str(p.stato),lat:typeof p.gpsY==='number'&&Number.isFinite(p.gpsY)?p.gpsY:null,lon:typeof p.gpsX==='number'&&Number.isFinite(p.gpsX)?p.gpsX:null};
  }));
  return {company:access.company,plants:rows.filter(Boolean),nextCursor:snap.size>limit?grants.at(-1).id:null};
}
module.exports={setPlantGrant,listMyPlants};

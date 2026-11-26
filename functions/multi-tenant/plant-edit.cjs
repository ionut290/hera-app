'use strict';
const {createHash}=require('node:crypto');
const fail=code=>{throw Object.assign(Error(code),{code});};
function boundary(){if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9199')fail('forbidden');}
async function editPlant({tx,access,caller,payload,id}){
 boundary();
 if(!['ordinary','snow'].includes(payload.kind))fail('invalid-kind');
 const jobId=id(payload.commessaId),plantId=id(payload.recordId),requestId=id(payload.requestId);
 if(!/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(requestId))fail('invalid-request');
 const revision=payload.revision;
 if(!Number.isSafeInteger(revision)||revision<1)fail('invalid-revision');
 const data=payload.data,keys=['denominazione','idSap','comune','indirizzo','gpsY','gpsX'];
 if(!data||typeof data!=='object'||Array.isArray(data)||Object.keys(data).length!==keys.length||keys.some(k=>!Object.hasOwn(data,k)))fail('invalid-payload');
 const clean={};
 for(const [k,max] of [['denominazione',150],['idSap',80],['comune',100],['indirizzo',200]]){
  if(typeof data[k]!=='string')fail('invalid-field');const s=data[k].trim().replace(/\s+/g,' ');if(s.length>max||(k==='denominazione'&&!s))fail('invalid-field');clean[k]=s;
 }
 for(const [k,max] of [['gpsY',90],['gpsX',180]]){const v=data[k];if(v!==null&&(typeof v!=='number'||!Number.isFinite(v)||Math.abs(v)>max))fail('invalid-coordinate');clean[k]=v;}
 if((clean.gpsY===null)!==(clean.gpsX===null))fail('invalid-coordinate');
 const root=access.root,job=root.collection(payload.kind==='snow'?'neve_commesse':'commesse').doc(jobId),p=job.collection('impianti').doc(plantId),f=job.collection('impiantiFisici').doc(plantId),receipt=root.collection('tenantWriteReceipts').doc(caller.uid+'_'+requestId);
 const hash=createHash('sha256').update(JSON.stringify(['editPlant',payload.kind,jobId,plantId,revision,clean])).digest('hex');
 const [j,a,b,r,work,all]=await Promise.all([tx.get(job),tx.get(p),tx.get(f),tx.get(receipt),tx.get(job.collection('lavorazioni').limit(1)),tx.get(job.collection('impianti').limit(1001))]);
 if(r.exists){if(r.data().hash!==hash)fail('invalid-request-reuse');return r.data().result;}
 if(!j.exists||j.data().tenantLabManaged!==true||!a.exists||!b.exists||a.data().tenantLabPlant!==true||b.data().tenantLabPlant!==true||a.data().physicalPlantId!==plantId||a.data().commessaId!==jobId||b.data().commessaId!==jobId)fail('invalid-reference');
 if(!work.empty||a.data().done!==false||a.data().stato!=='DA FARE'||b.data().statoGenerale!=='DA FARE')fail('invalid-linked-plant');
 if(a.data().tenantRevision!==revision||b.data().tenantRevision!==revision)fail('invalid-revision');
 if(all.size>1000)fail('invalid-plant-limit');
 const key=x=>JSON.stringify([x.idSap,x.denominazione,x.comune,x.indirizzo].map(v=>String(v||'').trim().toLowerCase()));
 if(all.docs.some(x=>x.id!==plantId&&(key(x.data())===key(clean)||(clean.idSap&&String(x.data().idSap||'').toLowerCase()===clean.idSap.toLowerCase()))))fail('already-exists');
 const meta={tenantRevision:revision+1,updatedAt:new Date(),updatedBy:caller.uid};
 tx.update(p,{...clean,...meta});
 tx.update(f,{denominazione:clean.denominazione,idSap:clean.idSap,comune:clean.comune,indirizzo:clean.indirizzo,latitudine:clean.gpsY,longitudine:clean.gpsX,...meta});
 const result={company:access.company,recordId:plantId,revision:revision+1};
 tx.create(receipt,{hash,result,createdAt:new Date()});
 tx.create(root.collection('audit').doc(),{event:'lab-plant-edited',recordId:plantId,commessaId:jobId,actorUid:caller.uid,at:new Date()});
 return result;
}
module.exports={editPlant};

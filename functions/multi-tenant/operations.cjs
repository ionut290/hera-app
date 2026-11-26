'use strict';
// Deliberately limited to synthetic emulator records; not the production operational runtime.
const {randomUUID,createHash}=require('node:crypto');
const digest=x=>createHash('sha256').update(JSON.stringify(x)).digest('hex');
function fail(code){throw Object.assign(new Error(code),{code});}
function txt(x,max=150,required=false){if(typeof x!=='string')fail('invalid-field');x=x.trim().replace(/\s+/g,' ');if(x.length>max||(required&&!x))fail('invalid-field');return x;}
function date(x){if(typeof x!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(x))fail('invalid-date');const d=new Date(x+'T00:00:00Z');if(!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==x)fail('invalid-date');return x;}
function ids(x,id,limit){if(!Array.isArray(x)||x.length>limit)fail('invalid-field');const list=x.map(id);if(new Set(list).size!==list.length)fail('invalid-duplicate');return list;}
function coords(x,max){if(x===null||x===undefined)return null;if(typeof x!=='number'||!Number.isFinite(x)||Math.abs(x)>max)fail('invalid-coordinate');return x;}
function emulatorOnly(){if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9199')fail('forbidden');}
async function createOperation({tx,access,caller,payload,id}){
  emulatorOnly();
  const operation=payload.operation,kind=payload.kind;
  if(!['plant','team','hoursRequest'].includes(operation))fail('invalid-operation');
  if(!['ordinary','snow'].includes(kind))fail('invalid-kind');
  const snow=kind==='snow',jobId=id(payload.commessaId),requestId=id(payload.requestId);
  if(!/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(requestId))fail('invalid-request');
  const data=payload.data,root=access.root;
  if(!data||typeof data!=='object'||Array.isArray(data))fail('invalid-payload');
  const allowed={plant:['denominazione','idSap','comune','indirizzo','gpsY','gpsX'],team:['date','leaderId','personnelIds','vehicleIds','plantIds'],hoursRequest:['date','personnelId','hours']}[operation];
  if(Object.keys(data).some(k=>!allowed.includes(k)))fail('invalid-payload');
  let clean;
  if(operation==='plant'){
    clean={denominazione:txt(data.denominazione,150,true),idSap:txt(data.idSap??'',80),comune:txt(data.comune??'',100),indirizzo:txt(data.indirizzo??'',200),gpsY:coords(data.gpsY,90),gpsX:coords(data.gpsX,180)};
    if((clean.gpsY===null)!==(clean.gpsX===null))fail('invalid-coordinate');
  }else if(operation==='team'){
    clean={date:date(data.date),leaderId:id(data.leaderId),personnelIds:ids(data.personnelIds,id,20),vehicleIds:ids(data.vehicleIds,id,10),plantIds:ids(data.plantIds,id,20)};
    if(clean.personnelIds.includes(clean.leaderId))fail('invalid-duplicate');
  }else{
    if(typeof data.hours!=='number'||!Number.isFinite(data.hours)||data.hours<=0||data.hours>24||Math.round(data.hours*4)!==data.hours*4)fail('invalid-hours');
    clean={date:date(data.date),personnelId:id(data.personnelId),hours:data.hours};
  }
  const job=root.collection(snow?'neve_commesse':'commesse').doc(jobId);
  const receipt=root.collection('tenantWriteReceipts').doc(caller.uid+'_'+requestId);
  const hash=digest({operation,kind,jobId,clean});
  const [jobSnap,prior]=await Promise.all([tx.get(job),tx.get(receipt)]);
  if(!jobSnap.exists||jobSnap.data().tenantLabManaged!==true)fail('invalid-unmanaged-parent');
  if(prior.exists){if(prior.data().hash!==hash)fail('invalid-request-reuse');return prior.data().result;}
  const now=new Date(),recordId=randomUUID(),jobName=txt(jobSnap.data().nome||'Commessa',150,true);
  const result={company:access.company,operation,kind,commessaId:jobId,recordId};
  const personCollection=root.collection(snow?'neve_personale':'personale'),vehicleCollection=root.collection(snow?'neve_mezzi':'mezzi');
  async function getManaged(ref){const s=await tx.get(ref);if(!s.exists||s.data().tenantLabManaged!==true)fail('invalid-reference');return s.data();}
  const personName=d=>txt(d.fullName||[d.cognome,d.nome].filter(Boolean).join(' '),200,true);
  const vehicleName=d=>txt(d.nId||d.nome,100,true);
  if(operation==='plant'){
    const [existing,physical,work]=await Promise.all([tx.get(job.collection('impianti').limit(1001)),tx.get(job.collection('impiantiFisici').limit(1001)),tx.get(job.collection('lavorazioni').limit(1))]);
    if(existing.size>1000||physical.size>1000)fail('invalid-plant-limit');
    // No mixing of the legacy/manual pilot with pre-existing economic or operational records.
    if(!work.empty||existing.docs.some(d=>d.data().tenantLabPlant!==true)||physical.docs.some(d=>d.data().tenantLabPlant!==true))fail('invalid-unmanaged-plant');
    const key=d=>JSON.stringify([d.idSap||'',d.denominazione||'',d.comune||'',d.indirizzo||''].map(v=>String(v).trim().toLowerCase()));
    if(existing.docs.some(d=>key(d.data())===key(clean)||(clean.idSap&&String(d.data().idSap||'').toLowerCase()===clean.idSap.toLowerCase())))fail('already-exists');
    const number=Math.max(0,Number(jobSnap.data().nextImpiantoNumber||1)-1,...existing.docs.map(d=>Number(d.data().numeroProgressivo)||0))+1;
    const metadata={commessaId:jobId,tenantLabManaged:true,tenantLabPlant:true,createdAt:now,updatedAt:now,createdBy:caller.uid};
    tx.create(job.collection('impiantiFisici').doc(recordId),{...metadata,denominazione:clean.denominazione,idSap:clean.idSap,comune:clean.comune,indirizzo:clean.indirizzo,latitudine:clean.gpsY,longitudine:clean.gpsX,numeroProgressivoImpianto:number,statoGenerale:'DA FARE'});
    tx.create(job.collection('impianti').doc(recordId),{...metadata,...clean,physicalPlantId:recordId,numeroProgressivo:number,sortOrder:number,stato:'DA FARE',done:false,doneAt:null,doneBy:'',codicePrezzo:''});
    tx.update(job,{nextImpiantoNumber:number+1});
  }else if(operation==='team'){
    const people=[clean.leaderId,...clean.personnelIds],vehicles=clean.vehicleIds;
    const current=root.collection(snow?'neve_squadre':'squadreCommesse').doc(jobId);
    const history=root.collection(snow?'neve_squadre_storico':'squadreStorico').doc(clean.date+'__'+jobId);
    // Conservative lab policy: one composition per job/day; no overwrite, overrides or absence bypass.
    const [c,h,p,v,plants,absence]=await Promise.all([tx.get(current),tx.get(history),Promise.all(people.map(x=>getManaged(personCollection.doc(x)))),Promise.all(vehicles.map(x=>getManaged(vehicleCollection.doc(x)))),Promise.all(clean.plantIds.map(x=>getManaged(job.collection('impianti').doc(x)))),tx.get(root.collection('calendarEvents').limit(1))]);
    if(h.exists||c.exists)fail('already-exists');
    // Until the original absence workflow is integrated, fail closed when any calendar event exists.
    if(!absence.empty)fail('invalid-absence-review');
    const locks=[...people.map(x=>['person',x]),...vehicles.map(x=>['vehicle',x])].map(([type,x])=>root.collection('tenantLabTeamLocks').doc(digest([clean.date,type,x])));
    if((await Promise.all(locks.map(x=>tx.get(x)))).some(x=>x.exists))fail('invalid-assignment-conflict');
    const row={caposquadra:personName(p[0]),personale:p.slice(1).map(personName).join(', '),mezzi:v.map(vehicleName).join(', '),impianti:plants.map(x=>x.denominazione).join(', '),impiantiDettagli:clean.plantIds.map((x,i)=>({id:x,denominazione:plants[i].denominazione,idSap:plants[i].idSap||''})),note:'',personnelIds:people,vehicleIds:vehicles};
    const team={commessaId:jobId,commessaNome:jobName,riferimentoData:clean.date,dateKey:clean.date,squadre:[row],tenantLabManaged:true,updatedAt:now,updatedBy:caller.uid};
    tx.create(current,team);tx.create(history,team);
    for(const lock of locks)tx.create(lock,{date:clean.date,commessaId:jobId,kind,createdAt:now});
    result.recordId=current.id;
  }else{
    const person=await getManaged(personCollection.doc(clean.personnelId));
    // Restrict to lab-managed personnel; no approval or external delivery.
    const name=personName(person);
    const lock=root.collection('tenantLabHoursLocks').doc(digest([clean.date,jobId,clean.personnelId]));
    const [existing,requests,reports]=await Promise.all([tx.get(lock),tx.get(root.collection(snow?'neve_ore_richieste':'oreApprovalRequests').limit(501)),tx.get(root.collection(snow?'neve_ore':'oreReports').limit(501))]);
    if(existing.exists)fail('invalid-hours-duplicate');
    if(requests.size>500||reports.size>500||[...requests.docs,...reports.docs].some(d=>d.data().tenantLabHours!==true))fail('invalid-hours-review');
    const entry={commessaId:jobId,commessaName:jobName,note:'',rows:[{operatore:name,operatoreId:clean.personnelId,ore:clean.hours,uniqueKey:jobId+'_'+clean.date+'_'+clean.personnelId}]};
    const request={date:clean.date,commessaName:jobName,entries:[entry],status:'pending_level1',level1ApprovedBy:'',level1ApprovedAt:null,level2ApprovedBy:'',level2ApprovedAt:null,finalizedReportId:'',createdByUid:caller.uid,createdByEmail:caller.email,createdAt:now,tenantLabHours:true,tenantLabManaged:true};
    tx.create(root.collection(snow?'neve_ore_richieste':'oreApprovalRequests').doc(recordId),request);
    tx.create(lock,{date:clean.date,commessaId:jobId,personnelId:clean.personnelId,requestId:recordId,status:'pending_level1',createdAt:now});
  }
  tx.create(receipt,{hash,result,createdAt:now});tx.create(root.collection('audit').doc(),{event:'lab-operation-'+operation,recordId:result.recordId,commessaId:jobId,kind,actorUid:caller.uid,at:now});
  return result;
}
module.exports={createOperation};

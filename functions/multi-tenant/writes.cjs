'use strict';
// Laboratory mutations. Only explicitly allowed fields, no deletes or protected workflows.
const {randomUUID,createHash}=require('node:crypto');
const core=require('../../inrete-work-items-v2.js');
const MODULES=new Set(['jobs','personnel','vehicles','prices']);
const fields=Object.freeze({jobs:['nome','codice','parentCommessaId'],personnel:['cognome','nome'],vehicles:['nId','marca','modello'],prices:['codiceVoce','descrizione','unitaMisura','prezzoBase','percentualeRibasso']});
function fail(code){throw Object.assign(new Error(code),{code});}
function value(v,max=100,required=false){if(typeof v!=='string')fail('invalid-field');const s=v.trim().replace(/\s+/g,' ');if(s.length>max||(required&&!s))fail('invalid-field');return s;}
function num(v,max){if(typeof v!=='number'||!Number.isFinite(v)||v<0||v>max)fail('invalid-field');return v;}
function validate(payload,id){
  const module=payload.module,kind=payload.kind;
  if(!MODULES.has(module))fail('invalid-module');
  if(!['ordinary','snow'].includes(kind))fail('invalid-kind');
  const commessaId=module==='prices'?id(payload.commessaId):null;
  if(module!=='prices'&&payload.commessaId!=null)fail('invalid-id');
  return {module,kind,commessaId};
}
function allowedData(module,data,id){
  if(!data||typeof data!=='object'||Array.isArray(data)||Object.keys(data).some(k=>!fields[module].includes(k)))fail('invalid-payload');
  if(module==='jobs')return {nome:value(data.nome,100,true),codice:value(data.codice??'',60),parentCommessaId:data.parentCommessaId?id(data.parentCommessaId):null};
  if(module==='personnel'){const cognome=value(data.cognome,100,true),nome=value(data.nome,100,true);return {cognome,nome,fullName:cognome+' '+nome};}
  if(module==='vehicles'){const nId=value(data.nId,60,true);return {nId,nome:nId,marca:value(data.marca??'',80),modello:value(data.modello??'',80)};}
  const codiceVoce=core.normalizePriceCode(value(data.codiceVoce,60,true));
  return {codiceVoce,descrizione:value(data.descrizione,200,true),unitaMisura:value(data.unitaMisura,20,true),prezzoBase:num(data.prezzoBase,100000000),percentualeRibasso:data.percentualeRibasso==null?null:num(data.percentualeRibasso,1)};
}
async function resolve(tx,access,payload,id){
  const {module,kind,commessaId}=validate(payload,id),root=access.root,snow=kind==='snow';
  let collection,parent=null;
  if(module==='jobs')collection=root.collection(snow?'neve_commesse':'commesse');
  if(module==='personnel')collection=root.collection(snow?'neve_personale':'personale');
  if(module==='vehicles')collection=root.collection(snow?'neve_mezzi':'mezzi');
  if(module==='prices'){
    parent=root.collection(snow?'neve_commesse':'commesse').doc(commessaId);
    const snap=await tx.get(parent);if(!snap.exists)fail('not-found');
    if(snap.data().tenantLabManaged!==true)fail('invalid-unmanaged-parent');
    collection=parent.collection('prezziario');
  }
  return {module,kind,commessaId,collection,parent};
}
function editable(module,d){return Object.fromEntries(fields[module].map(k=>[k,d[k]??(k==='percentualeRibasso'?null:k==='prezzoBase'?0:'')]));}
async function getRecord({tx,access,payload,id}){
  const target=await resolve(tx,access,payload,id);
  const doc=await tx.get(target.collection.doc(id(payload.recordId)));
  if(!doc.exists||doc.data().tenantLabManaged!==true)fail('not-found');
  return {company:access.company,module:target.module,kind:target.kind,commessaId:target.commessaId,recordId:doc.id,revision:doc.data().tenantRevision,data:editable(target.module,doc.data())};
}
async function saveRecord({tx,access,caller,payload,id}){
  const target=await resolve(tx,access,payload,id),{module,kind,commessaId,collection,parent}=target;
  const requestId=id(payload.requestId);
  if(!/^[a-f0-9-]{36}$/i.test(requestId))fail('invalid-request');
  const recordId=payload.recordId==null?null:id(payload.recordId),revision=payload.revision??0;
  if(!Number.isSafeInteger(revision)||revision<0||(!recordId&&revision!==0))fail('invalid-revision');
  const clean=allowedData(module,payload.data,id);
  const hash=createHash('sha256').update(JSON.stringify({module,kind,commessaId,recordId,revision,clean})).digest('hex');
  const receipt=access.root.collection('tenantWriteReceipts').doc(caller.uid+'_'+requestId);
  const previous=await tx.get(receipt);
  if(previous.exists){const p=previous.data();if(p.hash!==hash)fail('invalid-request-reuse');return p.result;}
  const ref=collection.doc(recordId||randomUUID()),old=await tx.get(ref);
  if(recordId&&(!old.exists||old.data().tenantLabManaged!==true))fail('not-found');
  if(recordId&&old.data().tenantRevision!==revision)fail('invalid-revision');
  if(!recordId&&old.exists)fail('already-exists');
  if(module==='jobs'&&clean.parentCommessaId){
    if(clean.parentCommessaId===ref.id)fail('invalid-parent');
    const p=await tx.get(collection.doc(clean.parentCommessaId));
    if(!p.exists||p.data().parentCommessaId)fail('invalid-parent');
  }
  if(module==='jobs'&&recordId&&(old.data().parentCommessaId??null)!==clean.parentCommessaId)fail('invalid-parent');
  let extra={},view=null,viewRows=null,viewKey=null;
  if(module==='jobs'&&!recordId)extra={excelModelVersion:2,priceListVersion:2,workItemsModelVersion:2,percentualeRibassoGenerale:0.01,nextImpiantoNumber:1,excelModelActivatedAt:new Date(),excelModelActivatedBy:caller.email,creatoDa:caller.email};
  if(module==='personnel'&&!recordId)extra={telefono:'',email:'',mansione:'',note:'',abilitatoTutteCommesse:false,commesseAbilitate:[],corsi:{}};
  if(module==='vehicles'&&!recordId)extra={portataCarico:'',massaComplessivaKg:'',alimentazione:'',posti:0};
  if(module==='prices'){
    // Do not change any price list containing operational work: original propagation
    // and historic accounting are protected and not implemented in this laboratory.
    const [work,plants,prices,parentSnap]=await Promise.all([
      tx.get(parent.collection('lavorazioni').limit(1)),tx.get(parent.collection('impianti').limit(1)),
      tx.get(collection.limit(501)),tx.get(parent)
    ]);
    if(!work.empty||!plants.empty)fail('invalid-linked-price');
    if(prices.size>500)fail('invalid-price-limit');
    if(prices.docs.some(p=>p.id!==ref.id&&core.normalizePriceCode(p.data().codiceVoce)===clean.codiceVoce))fail('already-exists');
    const general=core.parseNumber(parentSnap.data().percentualeRibassoGenerale)??0;
    if(general<0||general>1)fail('invalid-discount');
    extra={commessaId,prezzoRibassato:core.calculateDiscountedPrice(clean.prezzoBase,general,clean.percentualeRibasso),numeroProgressivo:old.exists?old.data().numeroProgressivo:Math.max(0,...prices.docs.map(p=>Number(p.data().numeroProgressivo)||0))+1};
  }
  if(kind==='ordinary'&&['personnel','vehicles'].includes(module)){
    view=access.root.collection('sharedStaticViews').doc('registri__corrente');
    const snap=await tx.get(view);viewKey=module==='personnel'?'personale':'mezzi';
    const rows=snap.exists?snap.data()?.payload?.[viewKey]:null;
    if(Array.isArray(rows)){
      if(rows.length>10000)fail('invalid-registry-limit');
      viewRows=rows.filter(x=>x?.id!==ref.id).concat([{...old.data(),...extra,...clean,id:ref.id,tenantLabManaged:true,tenantRevision:revision+1}]);
    }
  }
  const patch={...extra,...clean,tenantLabManaged:true,tenantRevision:revision+1,updatedAt:new Date(),updatedBy:caller.uid};
  if(!recordId)tx.create(ref,{...patch,createdAt:new Date()});else tx.update(ref,patch);
  if(viewRows)tx.update(view,{['payload.'+viewKey]:viewRows,updatedAt:new Date()});
  const result={company:access.company,module,kind,commessaId,recordId:ref.id,revision:revision+1};
  tx.create(receipt,{hash,result,createdAt:new Date()});
  tx.create(access.root.collection('audit').doc(),{event:'lab-record-'+(recordId?'updated':'created'),module,kind,commessaId,recordId:ref.id,actorUid:caller.uid,at:new Date()});
  return result;
}
module.exports={getRecord,saveRecord};

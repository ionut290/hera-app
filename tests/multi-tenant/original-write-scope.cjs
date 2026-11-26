'use strict';
// Local-only bridge for two original creation handlers. Reuses the existing
// quota-enforcing server dispatcher; does not expose Admin SDK or a raw DB.
// Uses nested tenant paths via the dispatcher, NOT the flat-path experiment.
const {randomUUID}=require('node:crypto');
const denied=code=>{throw Object.assign(Error(code),{code});};
const id=x=>typeof x==='string'&&/^[A-Za-z0-9_-]{1,128}$/.test(x);
function createOriginalWriteScope({tenantId,uid,currentUid,dispatch,token}){
 if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9199'||!id(tenantId)||!id(uid)||typeof currentUid!=='function'||typeof dispatch!=='function'||typeof token!=='function')denied('emulator-only');
 let active=true;
 const guard=()=>{if(!active||currentUid()!==uid)denied('permission-denied');};
 const allowed=(data,keys)=>{if(!data||typeof data!=='object'||Array.isArray(data)||Object.keys(data).some(k=>!keys.includes(k)))denied('invalid-payload');};
 async function send(action,payload){
  guard();const credential=await token();guard();
  const result=await dispatch(credential,action,{...payload,tenantId,requestId:randomUUID()});
  guard();if(result.company?.id!==tenantId||!id(result.recordId))denied('invalid-result');
  return Object.freeze({id:result.recordId});
 }
 const db=Object.freeze({collection(name){
  guard();if(!['commesse','neve_commesse'].includes(name))denied('permission-denied');
  const kind=name==='commesse'?'ordinary':'snow';
  return Object.freeze({
   add(data){
    guard();
    allowed(data,['nome','codice','excelModelVersion','priceListVersion','workItemsModelVersion','percentualeRibassoGenerale','excelModelActivatedAt','excelModelActivatedBy','nextImpiantoNumber','parentCommessaId','creatoDa','createdAt']);
    if(data.excelModelVersion!==2||data.priceListVersion!==2||data.workItemsModelVersion!==2||data.percentualeRibassoGenerale!==0.01||data.nextImpiantoNumber!==1)denied('invalid-payload');
    return send('saveCompanyRecord',{module:'jobs',kind,data:{nome:data.nome,codice:data.codice,parentCommessaId:data.parentCommessaId}});
   },
   doc(commessaId){
    guard();if(!id(commessaId))denied('invalid-id');
    return Object.freeze({collection(child){
     guard();if(child!=='impianti')denied('permission-denied');
     return Object.freeze({add(data){
      guard();
      allowed(data,['denominazione','idSap','tipologiaImpianto','comune','indirizzo','gpsY','gpsX','codicePrezzo','area','dittaEsecutrice','note','stato','done','doneAt','doneBy','numeroProgressivo','commessaId','createdAt','updatedAt','updatedByUid','updatedByName']);
      if(data.commessaId!==commessaId||data.stato!=='DA FARE'||data.done!==false||data.doneAt!==null||data.doneBy!==''||data.tipologiaImpianto!==''||data.codicePrezzo!==''||data.area!==''||data.dittaEsecutrice!==''||data.note!=='')denied('invalid-payload');
      return send('createCompanyOperation',{operation:'plant',kind,commessaId,data:{denominazione:data.denominazione,idSap:data.idSap,comune:data.comune,indirizzo:data.indirizzo,gpsY:data.gpsY,gpsX:data.gpsX}});
     }});
    }});
   }
  });
 }});
 return Object.freeze({db,close(){active=false;}});
}
module.exports={createOriginalWriteScope};

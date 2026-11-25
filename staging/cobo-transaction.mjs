import {resolveTenantContext} from './tenant-context.mjs';
import {planCoboWrite} from './cobo-write-plan.mjs';
const fields=new Set('id commessaId idSap denominazione nome comune indirizzo descrizioneVia area competenza gpsY gpsX latitudine longitudine coordinate tipologia tipologiaImpianto tipologiaIntervento lavorazioniRichieste codicePrezzo voceRiferimento hasOrdinario hasStraordinario tipoManutenzione sfalci sfalciMq note noteImpianto noteOperatore hasNote sfalcioCobo parcoCobo parcoCodvia parcoNomevia parcoQuartiere parcoConfiniDisponibili parcoDettagliPrimiSei parcoFonte parcoDataset schedaVersione updatedByUid updatedByName'.split(' '));
// Not wired to original UI. Current deployed rules deny these writes.
export async function saveCoboDemo({db,user,tenantId,payload,sdk}){
 if(db?.app?.options?.projectId!=='varga-cantieri-v2-staging')throw Error('Staging required');
 if(!user?.uid||typeof tenantId!=='string'||!/^[a-z0-9-]{3,64}$/.test(tenantId))throw Error('Invalid request');
 if(!payload||Object.keys(payload).some(k=>!fields.has(k)))throw Error('Unexpected payload field');
 const token=await user.getIdTokenResult(true);
 const {doc,runTransaction,serverTimestamp}=sdk;
 return runTransaction(db,async tx=>{
 const read=async(...p)=>{const s=await tx.get(doc(db,...p));return s.exists()?s.data():null;};
 const [tenant,member,membership]=await Promise.all([read('tenants',tenantId),read('tenants',tenantId,'members',user.uid),read('users',user.uid,'memberships',tenantId)]);
 const context=resolveTenantContext({uid:user.uid,claims:token.claims,tenantId,tenant,member,membership});
 if(!context)throw Error('Access denied');
 const commessa=await read('tenants',tenantId,'commesse','sfalcio-cobo');
 const plan=planCoboWrite({context,commessaExists:!!commessa,plantId:payload.id,payload});
 const plant=await read(...plan.plantPath);
 if(plan.initialize)tx.set(doc(db,...plan.commessaPath),{nome:'Sfalcio COBO',codice:'COBO-SFALCIO',speciale:true,tipoSpeciale:'SFALCIO_COBO',verdeBologna:true,parchiGiardini:true,attiva:true,parentCommessaId:null,excelModelVersion:2,priceListVersion:2,workItemsModelVersion:2,percentualeRibassoGenerale:0.01,nextImpiantoNumber:1,creatoDa:user.email||'',createdAt:serverTimestamp()});
 tx.set(doc(db,...plan.plantPath),{...payload,updatedByUid:user.uid,updatedByName:user.displayName||user.email||'Operatore',updatedAt:serverTimestamp(),...(!plant?{stato:'da_fare',done:false,doneAt:null,doneBy:'',createdAt:serverTimestamp()}:{})},{merge:true});
 return {plantId:payload.id,created:!plant,alreadyDone:plant?.done===true};
 });
}

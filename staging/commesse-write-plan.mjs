import {commessaDocumentPath,commessePath} from './commesse-paths.mjs';
const safeId=v=>{if(typeof v!=='string'||!/^[a-zA-Z0-9_-]{1,100}$/.test(v))throw Error('Invalid ID');return v;};
const commessaFields=new Set(['nome','codice','descrizione','cliente','committente','comune','indirizzo','attiva','tipo','createdAt','updatedAt','createdBy','updatedBy']);
const plantFields=new Set(['id','commessaId','nome','denominazione','comune','indirizzo','descrizione','coordinate','latitudine','longitudine','gpsX','gpsY','tipologia','note','stato','createdAt','updatedAt','createdBy','updatedBy']);
export function planCreateCommessa({context,commessaId,data}){
 if(!context?.uid||!['admin','supervisor'].includes(context.role))throw Error('Administrator or supervisor required');
 safeId(commessaId);if(!data||Object.keys(data).some(k=>!commessaFields.has(k))||typeof data.nome!=='string'||!data.nome.trim())throw Error('Invalid commessa data');
 return Object.freeze({path:Object.freeze(commessaDocumentPath(context,commessaId)),data:Object.freeze({...data,createdBy:context.uid})});
}
export function planCreatePlant({context,commessaId,plantId,data}){
 if(!context?.uid||!['admin','supervisor','employee'].includes(context.role))throw Error('Tenant role required');
 safeId(commessaId);safeId(plantId);
 if(!data||Object.keys(data).some(k=>!plantFields.has(k))||data.id!==plantId||data.commessaId!==commessaId||typeof (data.nome||data.denominazione)!=='string')throw Error('Invalid plant data');
 for(const k of ['done','doneAt','doneBy','fatto','completato','completedAt','tenantId'])if(Object.hasOwn(data,k))throw Error('Protected field');
 return Object.freeze({path:Object.freeze(commessePath(context,commessaId,plantId)),data:Object.freeze({...data,createdBy:context.uid})});
}

import {commessePath,commessaDocumentPath} from './commesse-paths.mjs';
// Pure planning only; no database writes.
export function planCoboWrite({context,commessaExists,plantId,payload}){
 if(!context?.uid||!['admin','supervisor','employee'].includes(context.role))throw Error('Access denied');
 if(!commessaExists&&context.role!=='admin')throw Error('Administrator required to initialize COBO');
 if(typeof plantId!=='string'||!/^cobo-parco-[a-z0-9-]{1,100}$/.test(plantId))throw Error('Invalid COBO ID');
 if(!payload||payload.id!==plantId||payload.commessaId!=='sfalcio-cobo'||payload.sfalcioCobo!==true)throw Error('Invalid COBO payload');
 for(const k of ['done','doneAt','doneBy','stato','createdAt','updatedAt','tenantId'])if(Object.hasOwn(payload,k))throw Error('Protected field');
 return Object.freeze({initialize:!commessaExists,commessaPath:Object.freeze(commessaDocumentPath(context,'sfalcio-cobo')),plantPath:Object.freeze(commessePath(context,'sfalcio-cobo',plantId))});
}

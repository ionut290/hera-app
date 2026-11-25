import {tenantCollectionPath} from './tenant-context.mjs';
function id(value) {
 if(typeof value !== 'string' || !value.trim() || value.includes('/') || value==='.' || value==='..') throw new Error('Invalid document ID');
 return value;
}
export function commessePath(context, commessaId, impiantoId) {
 const [root,tenant]=tenantCollectionPath(context,'sites');
 if(!context.uid || !['admin','supervisor','employee'].includes(context.role)) throw new Error('Tenant role required');
 const path=[root,tenant,'commesse'];
 if(commessaId===undefined) {
  if(impiantoId!==undefined) throw new Error('Commessa required');
  return path;
 }
 path.push(id(commessaId),'impianti');
 if(impiantoId!==undefined) path.push(id(impiantoId));
 return path;
}
export function commessaDocumentPath(context,commessaId) {
 if(commessaId===undefined) throw new Error('Commessa required');
 return commessePath(context,commessaId).slice(0,-1);
}

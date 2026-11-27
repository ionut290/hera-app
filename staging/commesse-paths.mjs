function id(value) {
 if(typeof value !== 'string' || !value.trim() || value.includes('/') || value==='.' || value==='..') throw new Error('Invalid document ID');
 return value;
}
export function commessePath(context, commessaId, impiantoId) {
 if(!context?.uid || typeof context.tenantId!=='string' || !/^[a-z0-9-]{3,64}$/.test(context.tenantId) || !['admin','supervisor','employee'].includes(context.role)) throw new Error('Tenant role required');
 const path=['tenants',context.tenantId,'commesse'];
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

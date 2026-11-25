import { tenantCollectionPath } from './tenant-context.mjs';
const allowed = new Set(['verdeLevatoRecords','verdeLevatoCommesse','verdeLevatoConfig']);
export function createModularVerdeStore(db, context, sdk) {
 if (db?.app?.options?.projectId !== 'varga-cantieri-v2-staging') throw new Error('Staging Firebase required');
 const [root,tenant] = tenantCollectionPath(context,'sites');
 if (!context.uid || !['admin','supervisor','employee'].includes(context.role)) throw new Error('Tenant role required');
 const {collection,doc,getDocsFromServer,getDocFromServer}=sdk;
 return Object.freeze({collection(name){
  if (!allowed.has(name)) throw new Error('Collection not approved');
  return Object.freeze({
   get:()=>getDocsFromServer(collection(db,root,tenant,name)),
   doc(id){
    if(typeof id!=='string'||!id||id.includes('/')) throw new Error('Document ID required');
    return Object.freeze({async get(){const snap=await getDocFromServer(doc(db,root,tenant,name,id));return {exists:snap.exists(),data:()=>snap.data()};}});
   }
  });
 }});
}

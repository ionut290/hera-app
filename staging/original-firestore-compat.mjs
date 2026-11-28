// Narrow Firestore compat boundary. Not installed in app.js; server rules remain mandatory.
// Route original collection names (including the normal/snow selectors) to the authenticated company.
import {quotaKind} from './trial-plan.mjs';
export function createOriginalFirestoreCompat({db,router}){
 const projectId=db?.app?.options?.projectId;
 if(!projectId || projectId!==router?.projectId || (projectId!=='varga-cantieri-v2-staging' && !projectId.startsWith('demo-varga-')))throw Error('Project mismatch');
 const tenantId=router?.context?.tenantId;
 if(typeof tenantId!=='string'||!/^[a-z0-9-]{3,64}$/.test(tenantId)||typeof router.path!=='function'||typeof router.close!=='function')throw Error('Verified router required');
 const prefix='tenants/'+tenantId+'/';
 let closed=false;
 function resolve(path,document=false){
  if(closed)throw Error('Scope closed');
  if(typeof path!=='string'||!path.length||path.startsWith('/')||path.endsWith('/'))throw Error('Invalid archive path');
  const local=path.startsWith(prefix)?path.slice(prefix.length):path;
  const segments=local.split('/');
  if(segments.some(s=>!s||s==='.'||s==='..'||s.includes('\\')||/[\u0000-\u001f]/.test(s)))throw Error('Invalid archive path');
  if(segments[0]==='tenants'||segments[0]==='users'||(segments.length%2===0)!==document)throw Error('Invalid archive scope');
  const root=router.path(segments[0]);
  if(root!==prefix+segments[0])throw Error('Unexpected archive scope');
  return [root,...segments.slice(1)].join('/');
 }
 function checkRef(ref){
  if(!ref||ref.firestore!==db||typeof ref.path!=='string'||!ref.path.startsWith(prefix)||resolve(ref.path,true)!==ref.path)throw Error('Reference outside company scope');
  return ref;
 }
 function open(){if(closed)throw Error('Scope closed');router.path('commesse');}
 function writes(native,transaction=false){
  let failed=false;let committed=false;
  function guard(){open();if(failed||committed)throw Error('Operation closed');}
  function apply(method,ref,...args){try{guard();native[method](checkRef(ref),...args);return facade;}catch(e){failed=true;throw e;}}
  const facade=Object.freeze({set:(ref,...args)=>apply('set',ref,...args),update:(ref,...args)=>apply('update',ref,...args),delete(){failed=true;throw Error('Deletion not enabled in staging compatibility boundary');},...(transaction?{get(ref){guard();return native.get(checkRef(ref));}}:{async commit(){guard();committed=true;return native.commit();}})});
  return {facade,verify:guard};
 }
 // Return actual compat SDK references. Their native query, batch and transaction
 // behavior is preserved; references already issued are governed by server rules.
 return Object.freeze({
  async writeQuotaDocument(ref,data,options){
   open();checkRef(ref);
   const p=ref.path.slice(prefix.length),kind=quotaKind(p);
   if(!kind||!['trial','full'].includes(router.company?.plan))return options===undefined?ref.set(data):ref.set(data,options);
   const parts=p.split('/'),tenant=db.doc('tenants/'+tenantId);
   return db.runTransaction(async tx=>{
    open();const snapshot=await tx.get(ref);
    if(!snapshot.exists){
     const t=(await tx.get(tenant)).data();
     open();
     if(!t||!Number.isSafeInteger(t.usage?.[kind])||!Number.isSafeInteger(t.limits?.[kind])||t.usage[kind]>=t.limits[kind])throw Error('Limite piano raggiunto: '+kind);
     tx.update(tenant,{['usage.'+kind]:t.usage[kind]+1,lastQuotaWrite:{kind,archive:parts[0],id:parts.at(-1),parent:parts.length===4?parts[1]:''}});
    }
    open();if(options===undefined)tx.set(ref,data);else tx.set(ref,data,options);
   });
  },
  collection(path){return db.collection(resolve(path));},
  doc(path){return db.doc(resolve(path,true));},
  batch(){open();return writes(db.batch()).facade;},
  runTransaction(fn,...options){open();if(typeof fn!=='function')throw Error('Transaction callback required');return db.runTransaction(async native=>{open();const scoped=writes(native,true);const result=await fn(scoped.facade);scoped.verify();return result;},...options);},
  close(){closed=true;router.close();}
 });
}

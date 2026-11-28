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
  const defer=transaction&&['trial','full'].includes(router.company?.plan),queued=[];
  function guard(){open();if(failed||committed)throw Error('Operation closed');}
  function apply(method,ref,...args){try{guard();checkRef(ref);if(defer)queued.push({method,ref,args});else native[method](ref,...args);return facade;}catch(e){failed=true;throw e;}}
  const facade=Object.freeze({set:(ref,...args)=>apply('set',ref,...args),update:(ref,...args)=>apply('update',ref,...args),delete(){failed=true;throw Error('Deletion not enabled in staging compatibility boundary');},...(transaction?{get(ref){guard();if(queued.length)throw Error('Transaction reads must precede writes');return native.get(checkRef(ref));}}:{async commit(){guard();committed=true;return native.commit();}})});
  async function finish(){
   guard();if(!defer)return;
   const archiveOf=op=>op.ref.path.slice(prefix.length).split('/')[0];
   const squadOps=queued.filter(op=>['squadreCommesse','neve_squadre','squadreStorico','neve_squadre_storico'].includes(archiveOf(op)));
   if(squadOps.length){
    const current=squadOps.find(op=>['squadreCommesse','neve_squadre'].includes(archiveOf(op)));
    if(queued.length!==2||squadOps.length!==2||!current||squadOps.some(op=>op.method!=='set'))throw Error('Salvataggio squadre atomico richiesto');
    const history=squadOps.find(op=>op!==current),payload=current.args[0],archive=archiveOf(current),id=current.ref.id;
    const historyArchive=archive==='squadreCommesse'?'squadreStorico':'neve_squadre_storico';
    if(!Array.isArray(payload?.squadre)||!payload.squadre.length||payload.commessaId!==id||history.ref.path!==prefix+historyArchive+'/'+payload.riferimentoData+'__'+id||!payload.updatedAt)throw Error('Composizione squadre non valida');
    const tenant=db.doc('tenants/'+tenantId);
    const [ts,cs]=await Promise.all([native.get(tenant),native.get(current.ref)]);
    guard();const t=ts.data(),old=cs.exists?cs.data().squadre:[];
    if(!Array.isArray(old)||!Number.isSafeInteger(t?.usage?.squadre)||!Number.isSafeInteger(t?.limits?.squadre))throw Error('Conteggio squadre da verificare');
    const count=t.usage.squadre-old.length+payload.squadre.length;
    if(count<0||count>t.limits.squadre)throw Error('Limite squadre operative raggiunto');
    native.update(tenant,{'usage.squadre':count,lastSquadraWrite:{archive,id,date:payload.riferimentoData,at:payload.updatedAt}});
   }
   for(const op of queued){guard();native[op.method](op.ref,...op.args);}
  }
  return {facade,verify:guard,finish};
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
  runTransaction(fn,...options){open();if(typeof fn!=='function')throw Error('Transaction callback required');return db.runTransaction(async native=>{open();const scoped=writes(native,true);const result=await fn(scoped.facade);await scoped.finish();scoped.verify();return result;},...options);},
  close(){closed=true;router.close();}
 });
}

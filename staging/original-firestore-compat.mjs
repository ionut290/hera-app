// Narrow Firestore compat boundary. Not installed in app.js; server rules remain mandatory.
// Route original collection names (including the normal/snow selectors) to the authenticated company.
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
 // Return actual compat SDK references. Their native query, batch and transaction
 // behavior is preserved; references already issued are governed by server rules.
 return Object.freeze({
  collection(path){return db.collection(resolve(path));},
  doc(path){return db.doc(resolve(path,true));},
  close(){closed=true;router.close();}
 });
}

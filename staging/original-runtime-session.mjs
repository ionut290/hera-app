// Entry boundary for the complete legacy app, not a replacement operational UI.
// Do not install in the production app or pass the raw SDK to legacy modules.
import {createOriginalArchiveRouter} from './original-archive-router.mjs';
import {createOriginalFirestoreCompat} from './original-firestore-compat.mjs';
const PROJECT='varga-cantieri-v2-staging';
export function createOriginalRuntimeSession({auth,db}){
 if(auth?.app?.options?.projectId!==PROJECT||db?.app?.options?.projectId!==PROJECT||auth.app!==db.app||typeof auth.onAuthStateChanged!=='function')throw Error('Staging SDK required');
 let revision=0,active=null,disposed=false;
 const current=()=>auth.currentUser;
 const invalidate=()=>{revision++;active?.firestore.close();active=null;};
 const unsubscribe=auth.onAuthStateChanged(user=>{if(active&&user?.uid!==active.context.uid)invalidate();},invalidate);
 // Read verified membership directly from server using the same compat SDK as the original app.
 const sdk={doc:(_db,...segments)=>db.doc(segments.join('/')),getDocFromServer:ref=>ref.get({source:'server'})};
 async function select(tenantId){
  if(disposed)throw Error('Runtime closed');
  invalidate();const request=revision,user=current();
  if(!user?.uid)throw Error('Active session required');
  const router=await createOriginalArchiveRouter({db,user,tenantId,sdk:{...sdk,getDocFromServer:async ref=>{
   const snap=await sdk.getDocFromServer(ref);
   return {exists:()=>snap.exists,data:()=>snap.data()};
  }},getCurrentUser:current});
  if(disposed||request!==revision||current()?.uid!==user.uid){router.close();throw Error('Selection cancelled');}
  const firestore=createOriginalFirestoreCompat({db,router});
  active=Object.freeze({company:router.company,context:router.context,firestore});
  return active;
 }
 return Object.freeze({select,clear:invalidate,close(){if(disposed)return;disposed=true;invalidate();unsubscribe();}});
}

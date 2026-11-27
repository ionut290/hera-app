import {createOriginalArchiveRouter} from './original-archive-router.mjs';
import {createOriginalCollectionBridge} from './original-collection-bridge.mjs';
// One active company per portal. Stale asynchronous opens cannot restore old selection.
export function createCompanySelection({db,sdk,getCurrentUser}){
 let revision=0,active=null;
 function close(){revision++;active?.close();active=null;}
 async function open(tenantId){
  close();const request=revision,user=getCurrentUser();
  if(!user?.uid)throw Error('Active session required');
  const router=await createOriginalArchiveRouter({db,sdk,user,tenantId,getCurrentUser});
  if(request!==revision||getCurrentUser()?.uid!==user.uid){router.close();throw Error('Selection cancelled');}
  const bridge=createOriginalCollectionBridge(router);
  active=Object.freeze({company:router.company,context:router.context,bridge,close:()=>router.close()});
  return active;
 }
 return Object.freeze({open,close});
}

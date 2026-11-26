'use strict';
// Narrow compat-shaped bridge for the unchanged original loader. Emulator only.
const {collection,query,orderBy,getDocsFromServer,onSnapshot}=require('firebase/firestore');
const denied=()=>Object.assign(new Error('Tenant scope unavailable'),{code:'permission-denied'});
const valid=x=>typeof x==='string'&&/^[A-Za-z0-9_-]{1,128}$/.test(x)&&!/^__.*__$/.test(x);
function createLegacyCommesseScope({firestore,tenantId,uid,currentUid}){
  if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||firestore?.app?.options?.projectId!=='demo-varga-multi-tenant'||!valid(tenantId)||!valid(uid)||typeof currentUid!=='function')throw denied();
  let active=true;const listeners=new Set();
  const guard=()=>{if(!active||currentUid()!==uid)throw denied();};
  function close(){active=false;for(const stop of [...listeners])stop();listeners.clear();}
  const db=Object.freeze({collection(name){
    guard();if(!['commesse','neve_commesse'].includes(name))throw denied();
    const scoped=collection(firestore,'tenants',tenantId,name);
    return Object.freeze({orderBy(field,direction){
      guard();if(field!=='createdAt'||direction!=='desc')throw denied();
      const q=query(scoped,orderBy(field,direction));
      return Object.freeze({
        async get(){guard();const result=await getDocsFromServer(q);guard();return result;},
        onSnapshot(next,error){
          guard();let live=true,unsubscribe;
          const stop=()=>{live=false;listeners.delete(stop);if(unsubscribe)unsubscribe();};
          listeners.add(stop);
          unsubscribe=onSnapshot(q,{includeMetadataChanges:true},snap=>{
            if(!live||!active||currentUid()!==uid){stop();return;}
            if(!snap.metadata.fromCache&&!snap.metadata.hasPendingWrites)next(snap);
          },err=>{if(live&&active&&currentUid()===uid&&typeof error==='function')error(err);stop();});
          return stop;
        }
      });
    }});
  }});
  return Object.freeze({db,close});
}
module.exports={createLegacyCommesseScope};

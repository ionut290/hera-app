'use strict';
// Read-only original commesse loader adapter. No protected FATTO code is called.
// Flat collection names retained; the query is narrowed only at this lab boundary.
const {collection,query,where,orderBy,getDocsFromServer,onSnapshot}=require('firebase/firestore');
const denied=()=>Object.assign(new Error('Tenant session unavailable'),{code:'permission-denied'});
const valid=x=>typeof x==='string'&&/^[A-Za-z0-9_-]{1,80}$/.test(x);
function createFlatSessionScope({firestore,tenantId,uid,sessionId,currentIdentity}){
 if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||firestore?.app?.options?.projectId!=='demo-varga-multi-tenant'||!valid(tenantId)||!valid(uid)||!valid(sessionId)||typeof currentIdentity!=='function')throw denied();
 let active=true;const listeners=new Set();
 const guard=()=>{const identity=currentIdentity();if(!active||identity?.uid!==uid||identity?.sessionId!==sessionId||identity?.tenantId!==tenantId)throw denied();};
 function close(){active=false;for(const stop of [...listeners])stop();listeners.clear();}
 const db=Object.freeze({collection(name){
  guard();if(!['commesse','neve_commesse'].includes(name))throw denied();
  return Object.freeze({orderBy(field,direction){
   guard();if(field!=='createdAt'||direction!=='desc')throw denied();
   const q=query(collection(firestore,name),where('tenantId','==',tenantId),where('flatSessionExperiment','==',true),orderBy(field,direction));
   return Object.freeze({
    async get(){guard();const result=await getDocsFromServer(q);guard();return result;},
    onSnapshot(next,error){
     guard();let live=true,unsubscribe;
     const stop=()=>{if(!live)return;live=false;listeners.delete(stop);if(unsubscribe)unsubscribe();};
     listeners.add(stop);
     try{unsubscribe=onSnapshot(q,{includeMetadataChanges:true},snap=>{
      try{guard();}catch{stop();return;}
      if(live&&!snap.metadata.fromCache&&!snap.metadata.hasPendingWrites)next(snap);
     },err=>{try{guard();if(live&&typeof error==='function')error(err);}catch{}finally{stop();}});if(!live)unsubscribe();}
     catch(e){stop();throw e;}
     return stop;
    }
   });
  }});
 }});
 return Object.freeze({db,close});
}
module.exports={createFlatSessionScope};

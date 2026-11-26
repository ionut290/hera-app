// Candidate bootstrap: verified company membership, never legacy global admin emails.
// Do not install this into the production app. No persistent browser/native data.
import {createOriginalRuntimeSession} from './original-runtime-session.mjs';
import {createOriginalArchiveRouter} from './original-archive-router.mjs';
export function createEphemeralRuntimeStorage({maxBytes=8*1024*1024}={}){
 if(!Number.isSafeInteger(maxBytes)||maxBytes<1)throw Error('Invalid storage quota');
 let closed=false,bytes=0;const entries=new Map();
 const check=()=>{if(closed)throw Error('Company storage closed');};
 const cost=(k,v)=>2*(k.length+v.length);
 const storage=Object.freeze({
  get length(){check();return entries.size;},
  key(index){check();return Number.isInteger(index)&&index>=0?[...entries.keys()][index]??null:null;},
  getItem(key){check();return entries.get(String(key))??null;},
  setItem(key,value){check();const k=String(key),v=String(value),old=entries.get(k);const next=bytes-(old===undefined?0:cost(k,old))+cost(k,v);if(next>maxBytes)throw Error('Company storage quota exceeded');entries.set(k,v);bytes=next;},
  removeItem(key){check();const k=String(key),old=entries.get(k);if(old!==undefined){bytes-=cost(k,old);entries.delete(k);}},
  clear(){check();entries.clear();bytes=0;}
 });
 return Object.freeze({storage,close(){if(closed)return;entries.clear();bytes=0;closed=true;}});
}
export function createOriginalRuntimeEntry({auth,db}){
 const runtime=createOriginalRuntimeSession({auth,db});
 let revision=0,active=null,closed=false;
 const current=()=>auth.currentUser;
 const invalidate=()=>{revision++;active?.local.close();active?.session.close();active=null;runtime.clear();};
 const unsubscribe=auth.onAuthStateChanged(user=>{if(active&&user?.uid!==active.scope.context.uid)invalidate();},invalidate);
 const sdk={doc:(_db,...p)=>db.doc(p.join('/')),getDocFromServer:async ref=>{const snapshot=await ref.get({source:'server'});return {exists:()=>snapshot.exists,data:()=>snapshot.data()};}};
 async function enter(tenantId){
  if(closed)throw Error('Entry closed');
  invalidate();const request=revision,scope=await runtime.select(tenantId);
  if(closed||request!==revision||current()?.uid!==scope.context.uid){scope.firestore.close();throw Error('Entry cancelled');}
  const local=createEphemeralRuntimeStorage(),session=createEphemeralRuntimeStorage();
  const selected={scope,local,session};active=selected;
  const live=()=>!closed&&active===selected&&current()?.uid===scope.context.uid;
  async function verify(user){
   if(!live()||user?.uid!==scope.context.uid)return {allowed:false};
   let fresh;
   try{
    fresh=await createOriginalArchiveRouter({db,user:current(),tenantId:scope.context.tenantId,sdk,getCurrentUser:current});
    if(!live()||fresh.context.role!==scope.context.role){if(active===selected)invalidate();return {allowed:false};}
    const role=fresh.context.role,admin=role==='admin';
    // Compatibility profile derived only from verified membership and the authenticated user.
    // Never reuse a global admin list or carry forward legacy profile privileges.
    const profile=Object.freeze({id:user.uid,uid:user.uid,email:current().email||'',displayName:current().displayName||current().email||'Utente',role:admin?'admin':'user',ruolo:admin?'admin':'user',tenantRole:role,isAdmin:admin,admin,statoAccount:'attivo',accountStatus:'attivo',accessApproved:true,banned:false,tenantId:scope.context.tenantId});
    return {allowed:true,status:'attivo',profile};
   }catch{if(active===selected)invalidate();return {allowed:false};}
   finally{fresh?.close();}
  }
  const approval=Object.freeze({verify,statusOf:profile=>profile?.statoAccount||'non autorizzato'});
  return Object.freeze({projectId:db.app.options.projectId,session:scope,auth,approval,localStorage:local.storage,sessionStorage:session.storage,isActive:live});
 }
 return Object.freeze({enter,leave:invalidate,close(){if(closed)return;closed=true;invalidate();runtime.close();unsubscribe();}});
}

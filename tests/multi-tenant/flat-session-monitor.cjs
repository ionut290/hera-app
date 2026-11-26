'use strict';
// Local-only session lifetime guard. This module is not installed in the app.
// Server authorization remains mandatory; cached data must be cleared on close.
const {doc,onSnapshot,getDocFromServer}=require('firebase/firestore');
const valid=x=>typeof x==='string'&&/^[A-Za-z0-9_-]{1,80}$/.test(x);
function createSessionMonitor({firestore,uid,sessionId,tenantId,onClose,clock=()=>Date.now(),pollMs=5000}){
 if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||firestore?.app?.options?.projectId!=='demo-varga-multi-tenant'||![uid,sessionId,tenantId].every(valid)||typeof onClose!=='function'||typeof clock!=='function'||!Number.isInteger(pollMs)||pollMs<100||pollMs>30000)throw Error('emulator-only');
 const ref=doc(firestore,'experimentSessions',uid,'devices',sessionId);
 let state='pending',unsubscribe=null,expiryTimer=null,pollTimer=null,deadline=0,flight=null;
 function close(reason='closed'){
  if(state==='closed')return;state='closed';
  if(unsubscribe)unsubscribe();clearTimeout(expiryTimer);clearInterval(pollTimer);deadline=0;
  onClose(reason);
 }
 function check(s){
  if(state==='closed')return false;
  if(!s.exists()||s.metadata.fromCache||s.metadata.hasPendingWrites){close('session-unverified');return false;}
  const d=s.data(),end=d.expiresAt?.toMillis?.();
  if(d.tenantId!==tenantId||d.status!=='active'||!Number.isFinite(end)||end<=clock()){close('session-invalid');return false;}
  deadline=end;clearTimeout(expiryTimer);
  expiryTimer=setTimeout(()=>close('session-expired'),Math.min(Math.max(0,end-clock()),2147483647));return true;
 }
 function verify(){
  if(state==='closed')return Promise.resolve(false);
  if(flight)return flight;
  const p=getDocFromServer(ref).then(check,()=>{close('session-unavailable');return false;});
  flight=p;return p.finally(()=>{if(flight===p)flight=null;});
 }
 async function start(){
  if(state!=='pending')throw Error('monitor already started');
  state='starting';
  if(!await verify()||state==='closed')return false;
  state='active';
  unsubscribe=onSnapshot(ref,{includeMetadataChanges:true},s=>{
   if(state==='closed')return;
   // A local/cache-only event is never used to open or extend a session.
   if(s.metadata.fromCache||s.metadata.hasPendingWrites)return;
   check(s);
  },()=>close('session-unavailable'));
  pollTimer=setInterval(()=>{void verify();},pollMs);
  return true;
 }
 function assertActive(){if(state!=='active'||clock()>=deadline){close('session-invalid');throw Object.assign(Error('Session unavailable'),{code:'permission-denied'});}}
 return Object.freeze({start,verify,assertActive,close,get state(){return state;}});
}
module.exports={createSessionMonitor};

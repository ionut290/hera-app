import firebase from 'firebase/compat/app';
import 'firebase/compat/auth';
import 'firebase/compat/firestore';
import {createOriginalRuntimeEntry} from '../original-runtime-entry.mjs';
import {createLegacyFirebaseNamespace} from '../legacy-firebase-namespace.mjs';
async function start(){
const $=id=>document.getElementById(id);
const projectId='varga-cantieri-v2-staging';
const status=m=>$('status').textContent=m;
const allowed=location.hostname==='127.0.0.1'||location.hostname==='localhost';
if(!allowed||!['http:','https:','capacitor:'].includes(location.protocol))throw Error('Local emulator only');
const app=firebase.initializeApp({apiKey:'demo-local-emulator-key',authDomain:'localhost',projectId,appId:'demo-local-emulator'});
const auth=app.auth(),db=app.firestore();
auth.useEmulator('http://127.0.0.1:9099',{disableWarnings:true});
db.useEmulator('127.0.0.1',8080);
await auth.setPersistence(firebase.auth.Auth.Persistence.NONE);
const entry=createOriginalRuntimeEntry({auth,db});
let prepared=null,revision=0;
function closeRuntime(){revision++;$('runtime').removeAttribute('src');$('runtime').hidden=true;prepared=null;entry.leave();}
function present(session){
 const statics={FieldValue:firebase.firestore.FieldValue,FieldPath:firebase.firestore.FieldPath,Timestamp:firebase.firestore.Timestamp,GeoPoint:firebase.firestore.GeoPoint};
 prepared=Object.freeze({session,namespace:createLegacyFirebaseNamespace({prepared:session,statics})});
 $('runtime').src='original.html';$('runtime').hidden=false;
}
// This is a local testing bridge, not a security boundary. The emulator rules are mandatory.
Object.defineProperty(window,'__vargaLocalRuntime',{value:()=>{if(!prepared?.session.isActive())throw Error('Verified local company required');return prepared;},writable:false,configurable:false});
$('login').addEventListener('submit',async e=>{e.preventDefault();status('Accesso in corso…');try{await auth.signInWithEmailAndPassword($('email').value,$('password').value);$('password').value='';}catch(err){status('Accesso non riuscito: '+err.code);}});
$('logout').onclick=async()=>{closeRuntime();await auth.signOut();};
$('back').onclick=()=>{closeRuntime();status('Seleziona un’altra azienda.');};
auth.onAuthStateChanged(async user=>{
 closeRuntime();const request=revision;$('companies').replaceChildren();$('login').hidden=Boolean(user);$('logout').hidden=!user;$('back').hidden=!user;
 if(!user){status('Accedi con un account di prova locale.');return;}
 try{
  const claims=(await user.getIdTokenResult(true)).claims;if(claims.email_verified!==true||claims.apple_review===true)throw Error('Account non autorizzato');
  const memberships=await db.collection('users').doc(user.uid).collection('memberships').get({source:'server'});
  if(request!==revision)return;
  for(const item of memberships.docs){
   if(item.data().status!=='active')continue;
   const tenant=await db.doc('tenants/'+item.id).get({source:'server'});
   if(request!==revision)return;
   if(!tenant.exists||tenant.data().demo!==true||tenant.data().status!=='active')continue;
   const button=document.createElement('button');button.textContent=tenant.data().name||item.id;button.dataset.tenant=item.id;
   button.onclick=async()=>{closeRuntime();const attempt=revision;status('Verifica azienda…');try{const session=await entry.enter(item.id);if(attempt!==revision)return;const approval=await session.approval.verify(auth.currentUser);if(!approval.allowed||attempt!==revision)throw Error('Accesso non autorizzato');present(session);status('Azienda locale: '+session.session.company.name);}catch(err){if(attempt===revision){closeRuntime();status('Accesso aziendale non disponibile.');}}};
   $('companies').append(button);
  }
  status('Seleziona azienda di prova.');
 }catch(err){status('Accesso non disponibile: '+(err.code||err.message));}
});
window.addEventListener('pagehide',()=>entry.close(),{once:true});

}
start().catch(e=>{document.getElementById('status').textContent='Avvio locale non disponibile: '+e.message;});

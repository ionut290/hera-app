import firebase from 'firebase/compat/app';
import 'firebase/compat/auth';
import 'firebase/compat/firestore';
import {createOriginalRuntimeEntry} from '../original-runtime-entry.mjs';
import {createLegacyFirebaseNamespace} from '../legacy-firebase-namespace.mjs';
import {connectLocalFirebase,connectStagingFirebase} from '../firebase-runtime-provider.mjs';
import {createAdminPanel} from './integrated-superadmin.js';
async function start(){
const $=id=>document.getElementById(id);
const status=m=>$('status').textContent=m;
const mode=__VARGA_BUILD_MODE__;
const isLive=mode==='private-staging';
if(!isLive&&mode!=='local-emulator')throw Error('Build mode not approved');
const apiKey=isLive?__VARGA_STAGING_WEB_KEY__:'demo-local-emulator-key';
const {auth,db,statics}=await (isLive?connectStagingFirebase({firebase,location,apiKey}):connectLocalFirebase({firebase,location}));
const entry=createOriginalRuntimeEntry({auth,db});
const adminPanel=createAdminPanel({host:$('admin-panel'),auth,db,statics,readOnly:isLive});
let prepared=null,revision=0,superAdmin=false,adminVisible=false;
function closeRuntime(){revision++;$('runtime').removeAttribute('src');$('runtime').hidden=true;$('company-return').hidden=true;document.body.classList.remove('company-runtime');prepared=null;entry.leave();}
function present(session){
 const namespace=createLegacyFirebaseNamespace({prepared:session,statics});
// Only the facade reaches the original frame; keep raw SDK in the launcher.
// The frame is still same-origin, so server rules remain the security boundary.
const view=Object.freeze({projectId:session.projectId,session:session.session,auth:namespace.auth(),approval:session.approval,localStorage:session.localStorage,sessionStorage:session.sessionStorage,isActive:session.isActive});
prepared=Object.freeze({session:view,namespace,apiKey});
 $('runtime').src='original.html';$('runtime').hidden=false;$('company-return').hidden=false;document.body.classList.add('company-runtime');
}
// This is a local testing bridge, not a security boundary. The emulator rules are mandatory.
Object.defineProperty(window,'__vargaLocalRuntime',{value:()=>{if(!prepared?.session.isActive())throw Error('Verified local company required');return prepared;},writable:false,configurable:false});
$('login').addEventListener('submit',async e=>{e.preventDefault();status('Accesso in corso…');try{await auth.signInWithEmailAndPassword($('email').value,$('password').value);$('password').value='';}catch(err){status('Accesso non riuscito: '+err.code);}});
$('logout').onclick=async()=>{closeRuntime();adminPanel.clear();await auth.signOut();};
$('back').onclick=()=>{closeRuntime();status('Seleziona un’altra azienda.');};
$('company-return').onclick=()=>{closeRuntime();status('Seleziona un’altra azienda.');};
$('admin-toggle').onclick=async()=>{
 if(!superAdmin||!auth.currentUser)return;
 closeRuntime();
 if(adminVisible){adminPanel.clear();adminVisible=false;$('admin-toggle').textContent='Super Admin';$('companies').hidden=false;status('Seleziona un’azienda con appartenenza verificata.');return;}
 $('companies').hidden=true;adminVisible=true;$('admin-toggle').textContent='Torna alle aziende';status('Area amministrativa.');await adminPanel.open(auth.currentUser);
};
auth.onAuthStateChanged(async user=>{
 closeRuntime();adminPanel.clear();superAdmin=false;adminVisible=false;const request=revision;$('companies').replaceChildren();$('companies').hidden=false;$('admin-toggle').hidden=true;$('admin-toggle').textContent='Super Admin';$('login').hidden=Boolean(user);$('logout').hidden=!user;$('back').hidden=!user;
 if(!user){status(isLive?'Accedi con un account staging verificato.':'Accedi con un account di prova locale.');return;}
 try{
  const claims=(await user.getIdTokenResult(true)).claims;if(claims.email_verified!==true||claims.apple_review===true)throw Error('Account non autorizzato');
  superAdmin=claims.superAdmin===true;$('admin-toggle').hidden=!superAdmin;
  const memberships=await db.collection('users').doc(user.uid).collection('memberships').get({source:'server'});
  if(request!==revision)return;
  for(const item of memberships.docs){
   if(item.data().status!=='active')continue;
   let tenant;
   try{tenant=await db.doc('tenants/'+item.id).get({source:'server'});}catch(err){if(err.code==='permission-denied'||err.code==='firestore/permission-denied')continue;throw err;}
   if(request!==revision)return;
   if(!tenant.exists||tenant.data().demo!==true||tenant.data().status!=='active')continue;
   const button=document.createElement('button');button.textContent=tenant.data().name||item.id;button.dataset.tenant=item.id;
   button.onclick=async()=>{closeRuntime();const attempt=revision;status('Verifica azienda…');try{const session=await entry.enter(item.id);if(attempt!==revision)return;const approval=await session.approval.verify(auth.currentUser);if(!approval.allowed||attempt!==revision)throw Error('Accesso non autorizzato');if(isLive){present(session);status('Azienda staging: '+session.session.company.name+' — '+session.session.context.role+'.');return;}present(session);status('Azienda locale: '+session.session.company.name);}catch(err){if(attempt===revision){closeRuntime();status('Accesso aziendale non disponibile.');}}};
   $('companies').append(button);
  }
  if(!$('companies').childElementCount){status(superAdmin?'Nessuna appartenenza aziendale operativa. Apri Super Admin per amministrare le aziende.':'Nessuna azienda autorizzata.');return;}
  status(isLive?'Seleziona azienda staging (accesso in sola consultazione delle appartenenze).':'Seleziona azienda di prova.');
 }catch(err){status('Accesso non disponibile: '+(err.code||err.message));}
});
window.addEventListener('pagehide',()=>{adminPanel.clear();entry.close();},{once:true});

}
start().catch(e=>{document.getElementById('status').textContent='Avvio locale non disponibile: '+e.message;});

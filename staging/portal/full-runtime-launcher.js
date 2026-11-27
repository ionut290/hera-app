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
function clearCompanyAdmin(){ $('company-admin-panel').hidden=true; $('company-admin-list').replaceChildren(); }
function showLogin(){ $('login-card').hidden=false; $('register-panel').hidden=true; $('join-panel').hidden=true; $('company-panel').hidden=true; clearCompanyAdmin(); }
function showRegister(){ $('login-card').hidden=true; $('register-panel').hidden=false; $('join-panel').hidden=true; $('company-panel').hidden=true; clearCompanyAdmin(); }
function showJoin(){ $('login-card').hidden=true; $('register-panel').hidden=true; $('join-panel').hidden=false; $('company-panel').hidden=true; clearCompanyAdmin(); }
function makeJoinCode(){return 'VC-'+crypto.randomUUID().replaceAll('-','').slice(0,8).toUpperCase();}
async function createCompany(name,node){
 const user=auth.currentUser;if(!user)throw Error('LOGIN_REQUIRED');
 await user.reload();const token=await user.getIdTokenResult(true);
 if(token.claims.email_verified!==true){try{await user.sendEmailVerification();}catch{}throw Error('EMAIL_NOT_VERIFIED');}
 const clean=String(name||'').trim();if(clean.length<2||clean.length>120)throw Error('COMPANY_NAME_INVALID');
 const tenantId='demo-'+crypto.randomUUID(),joinCode=makeJoinCode(),now=statics.FieldValue.serverTimestamp();
 const tenant=db.doc('tenants/'+tenantId),codeDoc=db.doc('tenantJoinCodes/'+joinCode),member=db.doc('tenants/'+tenantId+'/members/'+user.uid),membership=db.doc('users/'+user.uid+'/memberships/'+tenantId);
 const batch=db.batch();
 batch.set(tenant,{name:clean,status:'active',demo:true,ownerUid:user.uid,joinCode,plan:'trial',trial:true,trialStartedAt:now,limits:{commesse:1,impianti:10,admin:1,operatori:3},usage:{commesse:0,impianti:0,admin:1,operatori:0},originalCreateEnabled:true,originalManualEnabled:true,originalArchivesEnabled:true,commesseWriteEnabled:true,snowCreateEnabled:true,fattoMode:'configurable-demo',createdAt:now,createdBy:user.uid});
 batch.set(codeDoc,{tenantId,companyName:clean,status:'active',createdAt:now,createdBy:user.uid});
 batch.set(member,{uid:user.uid,email:user.email||'',displayName:user.displayName||user.email||'Amministratore',role:'admin',status:'active',demo:true,createdAt:now});
 batch.set(membership,{tenantId,role:'admin',status:'active',demo:true,createdAt:now});
 await batch.commit();localStorage.setItem('varga.pendingCompanyName','');node.textContent='Azienda registrata. Codice azienda: '+joinCode+'. Apertura dell’app completa…';location.reload();
}
async function requestJoin(code,node){
 const user=auth.currentUser;if(!user)throw Error('LOGIN_REQUIRED');
 await user.reload();const token=await user.getIdTokenResult(true);
 if(token.claims.email_verified!==true){try{await user.sendEmailVerification();}catch{}throw Error('EMAIL_NOT_VERIFIED');}
 const clean=String(code||'').trim().toUpperCase().replace(/\s/g,'');if(!/^VC-[A-Z0-9]{8}$/.test(clean))throw Error('CODE_INVALID');
 const codeSnap=await db.doc('tenantJoinCodes/'+clean).get({source:'server'});if(!codeSnap.exists||codeSnap.data().status!=='active')throw Error('CODE_NOT_FOUND');
 const tenantId=codeSnap.data().tenantId,tenant=await db.doc('tenants/'+tenantId).get({source:'server'});if(!tenant.exists||tenant.data().status!=='active')throw Error('COMPANY_NOT_AVAILABLE');
 const ref=db.doc('tenants/'+tenantId+'/joinRequests/'+user.uid),existing=await ref.get({source:'server'});
 if(existing.exists&&existing.data().status==='pending')throw Error('REQUEST_PENDING');
 await ref.set({uid:user.uid,email:user.email||'',displayName:user.displayName||user.email||'Utente',status:'pending',requestedAt:statics.FieldValue.serverTimestamp()});
 node.textContent='Richiesta inviata a '+tenant.data().name+'. Attendi l’approvazione dell’amministratore.';
}
async function showCompanyAdmin(tenantId,tenantData){
 clearCompanyAdmin();const panel=$('company-admin-panel'),list=$('company-admin-list'),feedback=$('company-admin-status');panel.hidden=false;
 const code=tenantData.joinCode||'non disponibile';feedback.textContent='Codice azienda: '+code+' — caricamento richieste…';
 try{
  const rows=await db.collection('tenants').doc(tenantId).collection('joinRequests').where('status','==','pending').get({source:'server'});list.replaceChildren();
  if(!rows.size){feedback.textContent='Codice azienda: '+code+' — nessuna richiesta in attesa.';return;}
  feedback.textContent='Codice azienda: '+code+' — richieste in attesa: '+rows.size;
  for(const item of rows.docs){
   const data=item.data(),card=document.createElement('article');card.className='request-card';
   const title=document.createElement('strong');title.textContent=data.displayName||data.email||item.id;
   const mail=document.createElement('span');mail.textContent=data.email||'';
   const actions=document.createElement('div');actions.className='request-actions';
   const accept=document.createElement('button');accept.textContent='Accetta';
   const reject=document.createElement('button');reject.textContent='Rifiuta';
   accept.onclick=async()=>{accept.disabled=reject.disabled=true;try{const now=statics.FieldValue.serverTimestamp(),batch=db.batch();batch.set(db.doc('tenants/'+tenantId+'/members/'+item.id),{uid:item.id,email:data.email||'',displayName:data.displayName||data.email||'Utente',role:'employee',status:'active',demo:true,createdAt:now});batch.set(db.doc('users/'+item.id+'/memberships/'+tenantId),{tenantId,role:'employee',status:'active',demo:true,createdAt:now});batch.update(item.ref,{status:'accepted',reviewedAt:now,reviewedBy:auth.currentUser.uid});await batch.commit();await showCompanyAdmin(tenantId,tenantData);}catch(err){feedback.textContent='Accettazione non riuscita: '+(err.code||err.message);accept.disabled=reject.disabled=false;}};
   reject.onclick=async()=>{reject.disabled=accept.disabled=true;try{await item.ref.update({status:'rejected',reviewedAt:statics.FieldValue.serverTimestamp(),reviewedBy:auth.currentUser.uid});await showCompanyAdmin(tenantId,tenantData);}catch(err){feedback.textContent='Rifiuto non riuscito: '+(err.code||err.message);accept.disabled=reject.disabled=false;}};
   actions.append(accept,reject);card.append(title,mail,actions);list.append(card);
  }
 }catch(err){feedback.textContent='Richieste non disponibili: '+(err.code||err.message);}
}
$('open-register').onclick=showRegister;
$('open-join').onclick=showJoin;
$('back-to-login').onclick=showLogin;
$('back-from-join').onclick=showLogin;
$('login').addEventListener('submit',async e=>{e.preventDefault();status('Accesso in corso…');try{await auth.signInWithEmailAndPassword($('email').value,$('password').value);$('password').value='';}catch(err){status('Accesso non riuscito: '+err.code);}});
$('register').addEventListener('submit',async e=>{e.preventDefault();const node=$('register-status');node.textContent='Creazione account…';try{const company=$('company-name').value.trim();const credential=await auth.createUserWithEmailAndPassword($('register-email').value.trim(),$('register-password').value);localStorage.setItem('varga.pendingCompanyName',company);await credential.user.sendEmailVerification();node.textContent='Account creato. Controlla l’email, verifica l’indirizzo e poi accedi per attivare la tua azienda.';await auth.signOut();}catch(err){node.textContent=err.code==='auth/email-already-in-use'?'Questa email è già registrata. Accedi e crea l’azienda.':'Registrazione non riuscita: '+(err.code||err.message);}});
$('company-form').addEventListener('submit',async e=>{e.preventDefault();const node=$('company-status');node.textContent='Attivazione piano prova…';try{await createCompany($('company-name-existing').value,node);}catch(err){node.textContent=err.message==='EMAIL_NOT_VERIFIED'?'Verifica prima l’email: ti abbiamo inviato un nuovo messaggio di verifica.':'Azienda non creata: '+(err.code||err.message);}});
$('join').addEventListener('submit',async e=>{e.preventDefault();const node=$('join-status');node.textContent='Invio richiesta…';try{await requestJoin($('join-code').value,node);}catch(err){node.textContent=err.message==='LOGIN_REQUIRED'?'Crea prima un account e accedi.':err.message==='EMAIL_NOT_VERIFIED'?'Verifica prima l’email: ti abbiamo inviato un nuovo messaggio di verifica.':err.message==='CODE_NOT_FOUND'?'Codice azienda non valido o non più attivo.':err.message==='REQUEST_PENDING'?'Richiesta già inviata e in attesa di risposta.':'Richiesta non inviata: '+(err.code||err.message);}});

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
 closeRuntime();adminPanel.clear();superAdmin=false;adminVisible=false;const request=revision;$('companies').replaceChildren();$('companies').hidden=false;$('admin-toggle').hidden=true;$('admin-toggle').textContent='Super Admin';$('login').hidden=Boolean(user);$('login-card').hidden=Boolean(user);$('register-panel').hidden=true;$('company-panel').hidden=!user;$('logout').hidden=!user;$('back').hidden=!user;
 if(!user){showLogin();status(isLive?'Accedi oppure crea la tua azienda.':'Accedi con un account di prova locale.');return;}
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
   button.onclick=async()=>{closeRuntime();const attempt=revision;status('Verifica azienda…');try{const session=await entry.enter(item.id);if(attempt!==revision)return;const approval=await session.approval.verify(auth.currentUser);if(!approval.allowed||attempt!==revision)throw Error('Accesso non autorizzato');if(isLive){present(session);status('Azienda staging caricata: '+session.session.company.name+' — '+session.session.context.role+'.');return;}present(session);status('Azienda locale: '+session.session.company.name);}catch(err){if(attempt===revision){closeRuntime();status('Accesso aziendale non disponibile.');}}};
   $('companies').append(button);
  }
  if(!$('companies').childElementCount){$('company-panel').hidden=superAdmin;const pending=localStorage.getItem('varga.pendingCompanyName');if(pending)$('company-name-existing').value=pending;status(superAdmin?'Nessuna appartenenza aziendale operativa. Apri Super Admin per amministrare le aziende.':'Nessuna azienda autorizzata: registra la tua azienda per iniziare.');return;}$('company-panel').hidden=true;
  status(isLive?'Seleziona azienda staging (accesso in sola consultazione delle appartenenze).':'Seleziona azienda di prova.');
 }catch(err){status('Accesso non disponibile: '+(err.code||err.message));}
});
window.addEventListener('pagehide',()=>{adminPanel.clear();entry.close();},{once:true});

}
start().catch(e=>{document.getElementById('status').textContent='Avvio locale non disponibile: '+e.message;});

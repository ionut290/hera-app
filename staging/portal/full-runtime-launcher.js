import firebase from 'firebase/compat/app';
import {TRIAL_DAYS,TRIAL_LIMITS,initialTrialUsage} from '../trial-plan.mjs';
import {SUBSCRIPTION_PLANS,subscriptionLabel,formatPrice} from '../subscription-plans.mjs';
import {nativeProviderBridge,getProviderCredential,socialIntent} from './social-auth.mjs';
import 'firebase/compat/auth';
import 'firebase/compat/firestore';
import {createOriginalRuntimeEntry} from '../original-runtime-entry.mjs';
import {createLegacyFirebaseNamespace} from '../legacy-firebase-namespace.mjs';
import {connectLocalFirebase,connectStagingFirebase} from '../firebase-runtime-provider.mjs';
import {createAdminPanel} from './integrated-superadmin.js';
import {createCompanyDocuments} from './company-documents.js';
async function start(){
const $=id=>document.getElementById(id);
const status=m=>{const node=auth.currentUser?$('session-status'):$('status');node.textContent=m;};
const mode=__VARGA_BUILD_MODE__;
const isLive=mode==='private-staging';
if(!isLive&&mode!=='local-emulator')throw Error('Build mode not approved');
const apiKey=isLive?__VARGA_STAGING_WEB_KEY__:'demo-local-emulator-key';
const {auth,db,statics}=await (isLive?connectStagingFirebase({firebase,location,apiKey}):connectLocalFirebase({firebase,location}));
const entry=createOriginalRuntimeEntry({auth,db});
const socialBridge=nativeProviderBridge(window.Capacitor,isLive&&__VARGA_SOCIAL_AUTH_READY__);
const adminPanel=createAdminPanel({host:$('admin-panel'),auth,db,statics,readOnly:isLive});
let documentPanel=null;let prepared=null,revision=0,superAdmin=false,adminVisible=false,registrationMode='company',creatingAccount=false,loginJoinCode='',activationBusy=false,verificationBusy=false;
const intentKey=uid=>'varga.onboarding.'+uid;
function saveIntent(uid,intent){localStorage.setItem(intentKey(uid),JSON.stringify(intent));}
function readIntent(uid){try{return JSON.parse(localStorage.getItem(intentKey(uid))||'null');}catch{return null;}}
function clearIntent(uid){localStorage.removeItem(intentKey(uid));}
function normalizeCode(value){const code=String(value||'').trim().toUpperCase().replace(/\s/g,'');if(!/^VC-[A-Z0-9]{8}$/.test(code))throw Error('CODE_INVALID');return code;}
function userError(err){return ({'auth/invalid-email':'Indirizzo email non valido.','auth/invalid-credential':'Email o password non corrette.','auth/wrong-password':'Email o password non corrette.','auth/user-not-found':'Email o password non corrette.','auth/too-many-requests':'Troppi tentativi. Riprova più tardi.','auth/network-request-failed':'Connessione non disponibile. Riprova.','auth/email-already-in-use':'Questa email è già registrata. Accedi con il tuo account.','auth/weak-password':'Scegli una password di almeno 8 caratteri.','permission-denied':'Operazione non autorizzata. Contatta l’amministratore.','firestore/permission-denied':'Operazione non autorizzata. Contatta l’amministratore.'})[err.code]||'Operazione non riuscita. Riprova o contatta l’assistenza.';}
// Tear down the old browsing context rather than queuing about:blank on it.
function replaceRuntimeFrame(src){
 const old=$('runtime'),frame=old.cloneNode(false);
 frame.removeAttribute('src');frame.hidden=!src;
 if(src)frame.setAttribute('src',src);
 old.replaceWith(frame);
}
function closeRuntime(){documentPanel?.close();documentPanel=null;revision++;prepared=null;entry.leave();if($('runtime').hasAttribute('src'))replaceRuntimeFrame(null);else $('runtime').hidden=true;$('company-return').hidden=true;document.body.classList.remove('company-runtime');}
function present(session){
 const namespace=createLegacyFirebaseNamespace({prepared:session,statics});
// Only the facade reaches the original frame; keep raw SDK in the launcher.
// The frame is still same-origin, so server rules remain the security boundary.
const view=Object.freeze({projectId:session.projectId,session:session.session,auth:namespace.auth(),approval:session.approval,localStorage:session.localStorage,sessionStorage:session.sessionStorage,isActive:session.isActive});
prepared=Object.freeze({session:view,namespace,apiKey});
 replaceRuntimeFrame('original.html');$('company-return').hidden=false;document.body.classList.add('company-runtime');
}
// This is a local testing bridge, not a security boundary. The emulator rules are mandatory.
Object.defineProperty(window,'__vargaOpenCompanyDocuments',{value:(plant=null)=>{if(isLive||!prepared?.session.isActive())return false;documentPanel?.close();const session=prepared.session;documentPanel=createCompanyDocuments({host:$('company-documents'),auth,tenantId:session.session.context.tenantId,plant,isActive:()=>prepared?.session===session&&session.isActive()});return true;},writable:false,configurable:false});
Object.defineProperty(window,'__vargaLocalRuntime',{value:()=>{if(!prepared?.session.isActive())throw Error('Verified local company required');return prepared;},writable:false,configurable:false});
function clearCompanyAdmin(){ $('company-admin-panel').hidden=true; $('company-admin-list').replaceChildren(); }
function showWelcome(){ $('verify-panel').hidden=true; $('login-card').hidden=false; $('auth-options').hidden=false; $('login').hidden=true; $('login-title').textContent='Benvenuto'; $('register-panel').hidden=true; $('join-panel').hidden=true; $('company-panel').hidden=true; $('session-status').hidden=true; $('status').textContent=''; loginJoinCode=''; clearCompanyAdmin(); }
function showLogin(code=''){ $('verify-panel').hidden=true;loginJoinCode=code; $('login-card').hidden=false; $('auth-options').hidden=true; $('login').hidden=false; $('login-title').textContent='Accedi'; $('register-panel').hidden=true; $('join-panel').hidden=true; $('company-panel').hidden=true; $('status').textContent=code?'Accedi per inviare la richiesta alla tua azienda.':''; clearCompanyAdmin(); }
function showRegister(mode='company'){ $('verify-panel').hidden=true;registrationMode=mode;const join=mode==='join'; $('login-card').hidden=true; $('register-panel').hidden=false; $('join-panel').hidden=true; $('company-panel').hidden=true; $('register-title').textContent=join?'Crea il tuo account':'Crea la tua azienda'; $('register-description').textContent=join?'Verifica l’email per inviare la richiesta di adesione.':'Registrati e verifica la tua email per attivare il piano prova.'; $('register-trial').hidden=join; $('company-name-field').hidden=join; $('company-name').required=!join; $('register-submit').textContent=join?'Crea account e continua':'Crea azienda gratuitamente'; $('register-status').textContent=''; clearCompanyAdmin(); }
function showJoin(){ $('verify-panel').hidden=true; $('login-card').hidden=true; $('register-panel').hidden=true; $('join-panel').hidden=false; $('company-panel').hidden=true; $('join-submit').textContent=auth.currentUser?'Invia richiesta':'Continua'; $('join-existing').hidden=Boolean(auth.currentUser); $('back-from-join').hidden=Boolean(auth.currentUser); $('join-status').textContent=''; clearCompanyAdmin(); }
function makeJoinCode(){return 'VC-'+crypto.randomUUID().replaceAll('-','').slice(0,8).toUpperCase();}
async function createCompany(name,node){
 if(activationBusy)throw Error('ACTIVATION_BUSY');activationBusy=true;
 try{
 const user=auth.currentUser;if(!user)throw Error('LOGIN_REQUIRED');
 await user.reload();const token=await user.getIdTokenResult(true);
 if(token.claims.email_verified!==true){try{await user.sendEmailVerification();}catch{}throw Error('EMAIL_NOT_VERIFIED');}
 const clean=String(name||'').trim();if(clean.length<2||clean.length>120)throw Error('COMPANY_NAME_INVALID');
 const previous=readIntent(user.uid);
 const retry=previous?.type==='company'&&previous.name===clean&&/^demo-[0-9a-f-]{36}$/.test(previous.tenantId||'')&&/^VC-[A-Z0-9]{8}$/.test(previous.joinCode||'');
 const tenantId=retry?previous.tenantId:'demo-'+crypto.randomUUID(),joinCode=retry?previous.joinCode:makeJoinCode();
 const ensureSession=()=>{if(auth.currentUser?.uid!==user.uid)throw Error('SESSION_CHANGED');};
 ensureSession();
 // Persist the same IDs before attempting a write, including manual activation.
 saveIntent(user.uid,{type:'company',name:clean,tenantId,joinCode});
 const now=statics.FieldValue.serverTimestamp();
 const tenant=db.doc('tenants/'+tenantId),codeDoc=db.doc('tenantJoinCodes/'+joinCode),member=db.doc('tenants/'+tenantId+'/members/'+user.uid),membership=db.doc('users/'+user.uid+'/memberships/'+tenantId);
 const alreadyCreated=async()=>{
  ensureSession();
  const existing=await membership.get({source:'server'});ensureSession();
  if(!existing.exists)return false;
  const m=existing.data(),t=await tenant.get({source:'server'});ensureSession();
  if(m.role!=='admin'||m.status!=='active'||!t.exists||t.data().ownerUid!==user.uid||t.data().joinCode!==joinCode)throw Error('ACTIVATION_CONFLICT');
  return true;
 };
 const finish=()=>{ensureSession();clearIntent(user.uid);node.textContent='Azienda registrata. Codice azienda: '+joinCode+'.';};
 if(await alreadyCreated()){finish();return;}
 ensureSession();
 const batch=db.batch();
 batch.set(tenant,{name:clean,status:'active',demo:true,ownerUid:user.uid,joinCode,plan:'trial',trial:true,trialStartedAt:now,trialDays:TRIAL_DAYS,limits:{...TRIAL_LIMITS},usage:initialTrialUsage(),originalCreateEnabled:true,originalManualEnabled:true,originalArchivesEnabled:true,commesseWriteEnabled:true,snowCreateEnabled:true,fattoMode:'configurable-demo',createdAt:now,createdBy:user.uid});
 batch.set(codeDoc,{tenantId,companyName:clean,status:'active',createdAt:now,createdBy:user.uid});
 batch.set(member,{uid:user.uid,email:user.email||'',displayName:user.displayName||user.email||'Amministratore',role:'admin',status:'active',demo:true,createdAt:now});
 batch.set(membership,{tenantId,role:'admin',status:'active',demo:true,createdAt:now});
 try{await batch.commit();}catch(err){
  // A lost response can mean a successful atomic commit. Verify server state;
  // never create a second company or automatically retry the write.
  let recovered=false;try{recovered=await alreadyCreated();}catch{}
  if(!recovered)throw err;
 }
 finish();
 }finally{activationBusy=false;}
}
async function requestJoin(code,node){
 const user=auth.currentUser;if(!user)throw Error('LOGIN_REQUIRED');
 await user.reload();const token=await user.getIdTokenResult(true);
 if(token.claims.email_verified!==true){try{await user.sendEmailVerification();}catch{}throw Error('EMAIL_NOT_VERIFIED');}
 const ensureSession=()=>{if(auth.currentUser?.uid!==user.uid)throw Error('SESSION_CHANGED');};ensureSession();
 const clean=normalizeCode(code);
 const codeSnap=await db.doc('tenantJoinCodes/'+clean).get({source:'server'});if(!codeSnap.exists||codeSnap.data().status!=='active')throw Error('CODE_NOT_FOUND');
 ensureSession();
 const tenantId=codeSnap.data().tenantId;const companyName=codeSnap.data().companyName||'azienda';
 const ref=db.doc('tenants/'+tenantId+'/joinRequests/'+user.uid),existing=await ref.get({source:'server'});
 ensureSession();
 if(existing.exists){const state=existing.data().status;node.textContent=state==='accepted'?'Richiesta approvata. Aggiorna le aziende per entrare.':state==='rejected'?'Richiesta rifiutata. Contatta l’amministratore.':'Richiesta già inviata a '+companyName+'. Attendi l’approvazione.';return state;}
 await ref.set({uid:user.uid,email:user.email||'',displayName:user.displayName||user.email||'Utente',status:'pending',joinCode:clean,requestedAt:statics.FieldValue.serverTimestamp()});
 ensureSession();node.textContent='Richiesta inviata a '+companyName+'. Attendi l’approvazione dell’amministratore.';return 'pending';
}
async function showCompanyAdmin(tenantId,tenantData){
 clearCompanyAdmin();const panel=$('company-admin-panel'),list=$('company-admin-list'),feedback=$('company-admin-status');panel.hidden=false;
 const code=tenantData.joinCode||'non disponibile';const adminUid=auth.currentUser?.uid;feedback.textContent='Codice azienda: '+code+' — caricamento richieste…';
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
   accept.onclick=async()=>{accept.disabled=reject.disabled=true;try{if(auth.currentUser?.uid!==adminUid)throw Error('SESSION_CHANGED');const now=statics.FieldValue.serverTimestamp(),batch=db.batch();batch.update(db.doc('tenants/'+tenantId),{'usage.operatori':statics.FieldValue.increment(1),lastOperatorUid:item.id});batch.set(db.doc('tenants/'+tenantId+'/members/'+item.id),{uid:item.id,email:data.email||'',displayName:data.displayName||data.email||'Utente',role:'employee',status:'active',demo:true,createdAt:now});batch.set(db.doc('users/'+item.id+'/memberships/'+tenantId),{tenantId,role:'employee',status:'active',demo:true,createdAt:now});batch.update(item.ref,{status:'accepted',reviewedAt:now,reviewedBy:auth.currentUser.uid});await batch.commit();await showCompanyAdmin(tenantId,tenantData);}catch(err){feedback.textContent=err.code==='permission-denied'?'Accettazione non consentita: verifica il limite degli operatori o chiedi lo sblocco al Super Admin.':'Accettazione non riuscita: '+(err.code||err.message);accept.disabled=reject.disabled=false;}};
   reject.onclick=async()=>{reject.disabled=accept.disabled=true;try{await item.ref.update({status:'rejected',reviewedAt:statics.FieldValue.serverTimestamp(),reviewedBy:auth.currentUser.uid});await showCompanyAdmin(tenantId,tenantData);}catch(err){feedback.textContent='Rifiuto non riuscito: '+(err.code||err.message);accept.disabled=reject.disabled=false;}};
   actions.append(accept,reject);card.append(title,mail,actions);list.append(card);
  }
 }catch(err){feedback.textContent='Richieste non disponibili: '+(err.code||err.message);}
}
$('open-register').onclick=()=>showRegister();
$('open-join').onclick=showJoin;
$('open-login').onclick=()=>showLogin();
$('back-to-login').onclick=()=>registrationMode==='join'?showJoin():showWelcome();
$('back-from-join').onclick=showWelcome;
$('back-from-login').onclick=()=>loginJoinCode?showJoin():showWelcome();
$('join-existing').onclick=()=>{try{showLogin(normalizeCode($('join-code').value));}catch{$('join-status').textContent='Inserisci un codice valido (VC- seguito da 8 caratteri).';}};
for(const button of document.querySelectorAll('[data-password-toggle]'))button.onclick=()=>{const input=$(button.dataset.passwordToggle),visible=input.type==='password';input.type=visible?'text':'password';button.textContent=visible?'Nascondi':'Mostra';button.setAttribute('aria-label',visible?'Nascondi password':'Mostra password');};
$('reset-password').onclick=async()=>{const email=$('email');if(!email.value||!email.checkValidity()){status('Inserisci prima il tuo indirizzo email.');email.focus();return;}try{await auth.sendPasswordResetEmail(email.value.trim());status('Se l’account esiste, riceverai un’email per reimpostare la password.');}catch(err){status(err.code==='auth/user-not-found'?'Se l’account esiste, riceverai un’email per reimpostare la password.':userError(err));}};
$('login').addEventListener('submit',async e=>{e.preventDefault();status('Accesso in corso…');const button=e.submitter;button.disabled=true;try{await auth.signInWithEmailAndPassword($('email').value.trim(),$('password').value);$('password').value='';}catch(err){status(userError(err));}finally{button.disabled=false;}});
$('register').addEventListener('submit',async e=>{
 e.preventDefault();const node=$('register-status'),button=$('register-submit'),company=$('company-name').value.trim(),name=$('register-name').value.trim(),email=$('register-email').value.trim(),password=$('register-password').value;
 if(password!==$('register-confirm').value){node.textContent='Le password non coincidono.';return;}
 if(name.length<2||(registrationMode==='company'&&company.length<2)){node.textContent='Completa correttamente i dati richiesti.';return;}
 let intent;try{intent=registrationMode==='join'?{type:'join',code:normalizeCode($('join-code').value)}:{type:'company',name:company};}catch{node.textContent='Codice aziendale non valido.';return;}
 creatingAccount=true;button.disabled=true;node.textContent='Creazione account…';
 try{
  const credential=await auth.createUserWithEmailAndPassword(email,password);saveIntent(credential.user.uid,intent);let feedback='';try{await credential.user.updateProfile({displayName:name});}catch{}try{await credential.user.sendEmailVerification();feedback='Email di verifica inviata.';}catch(err){feedback='Non è stato possibile inviare l’email. Usa il pulsante per riprovare.';}$('register-password').value=$('register-confirm').value='';creatingAccount=false;await refreshAccount(credential.user);if(!$('verify-panel').hidden)$('verify-status').textContent=feedback;
 }catch(err){node.textContent=userError(err);if(auth.currentUser){creatingAccount=false;await refreshAccount(auth.currentUser);}}
 finally{creatingAccount=false;button.disabled=false;}
});
$('company-form').addEventListener('submit',async e=>{e.preventDefault();const node=$('company-status');node.textContent='Attivazione piano prova…';try{await createCompany($('company-name-existing').value,node);await refreshAccount(auth.currentUser);}catch(err){node.textContent=err.message==='EMAIL_NOT_VERIFIED'?'Verifica prima l’email.':userError(err);}});
function joinError(err){return ({CODE_INVALID:'Inserisci un codice valido (VC- seguito da 8 caratteri).',EMAIL_NOT_VERIFIED:'Verifica prima l’email.',CODE_NOT_FOUND:'Codice aziendale non valido o non più attivo.'})[err.message]||userError(err);}
$('join').addEventListener('submit',async e=>{e.preventDefault();const node=$('join-status');try{const code=normalizeCode($('join-code').value);if(!auth.currentUser){showRegister('join');return;}const uid=auth.currentUser.uid;node.textContent='Invio richiesta…';await requestJoin(code,node);if(auth.currentUser?.uid!==uid)return;saveIntent(uid,{type:'join',code});}catch(err){node.textContent=joinError(err);}});
$('logout').onclick=async()=>{closeRuntime();adminPanel.clear();await auth.signOut();};
$('back').onclick=()=>{if(auth.currentUser&&!activationBusy&&!creatingAccount&&!verificationBusy)refreshAccount(auth.currentUser);};
$('company-return').onclick=()=>{closeRuntime();status('Seleziona un’altra azienda.');};
$('admin-toggle').onclick=async()=>{
 if(!superAdmin||!auth.currentUser)return;
 closeRuntime();
 if(adminVisible){adminPanel.clear();adminVisible=false;$('admin-toggle').textContent='Super Admin';$('companies').hidden=false;status('Seleziona un’azienda con appartenenza verificata.');return;}
 $('companies').hidden=true;adminVisible=true;$('admin-toggle').textContent='Torna alle aziende';status('Area amministrativa.');await adminPanel.open(auth.currentUser);
};
async function refreshAccount(user){
 if(creatingAccount&&user)return;
 $('verify-panel').hidden=true;$('verify-status').textContent='';
 closeRuntime();adminPanel.clear();superAdmin=false;adminVisible=false;const request=revision;let companyCount=0;$('companies').replaceChildren();$('companies').hidden=false;$('admin-toggle').hidden=true;$('admin-toggle').textContent='Super Admin';$('login-card').hidden=Boolean(user);$('session-status').hidden=!user;$('session-status').textContent='';$('join-existing').hidden=Boolean(user);$('back-from-join').hidden=Boolean(user);$('join-submit').textContent=user?'Invia richiesta':'Continua';$('register-panel').hidden=true;$('join-panel').hidden=true;$('company-panel').hidden=true;$('company-admin-panel').hidden=true;$('logout').hidden=!user;$('back').hidden=!user;
 if(!user){if(!creatingAccount)showWelcome();return;}
 try{
  const joinCode=loginJoinCode;loginJoinCode='';if(joinCode)saveIntent(user.uid,{type:'join',code:joinCode});
  await user.reload();const claims=(await user.getIdTokenResult(true)).claims;if(request!==revision||auth.currentUser?.uid!==user.uid)return;
  if(claims.apple_review===true){$('back').hidden=true;status('Account non autorizzato per questa versione.');return;}
  if(claims.email_verified!==true){$('back').hidden=true;$('verify-panel').hidden=false;$('verify-email').textContent=user.email||'';status('Accesso aziendale sospeso fino alla verifica dell’email.');return;}
  $('join-panel').hidden=false;
  superAdmin=claims.superAdmin===true;$('admin-toggle').hidden=!superAdmin;
  const memberships=await db.collection('users').doc(user.uid).collection('memberships').get({source:'server'});
  if(request!==revision)return;
  for(const item of memberships.docs){
   if(item.data().status!=='active')continue;
   let tenant;
   try{tenant=await db.doc('tenants/'+item.id).get({source:'server'});}catch(err){if(err.code==='permission-denied'||err.code==='firestore/permission-denied')continue;throw err;}
   if(request!==revision)return;
   if(!tenant.exists||tenant.data().demo!==true||tenant.data().status!=='active')continue;
   companyCount++;const button=document.createElement('button');const t=tenant.data();
   button.textContent=(t.name||item.id)+' · '+subscriptionLabel(t);button.dataset.tenant=item.id;
   button.onclick=async()=>{closeRuntime();const attempt=revision;status('Verifica azienda…');try{const session=await entry.enter(item.id);if(attempt!==revision)return;const approval=await session.approval.verify(auth.currentUser);if(!approval.allowed||attempt!==revision)throw Error('Accesso non autorizzato');if(isLive){present(session);status('Azienda caricata: '+session.session.company.name+' — '+session.session.context.role+'.');return;}present(session);status('Azienda locale: '+session.session.company.name);}catch(err){if(attempt===revision){closeRuntime();status('Accesso aziendale non disponibile.');}}};
   $('companies').append(button);if(item.data().role==='admin'){const manage=document.createElement('button');manage.textContent='Gestisci codice e richieste';manage.onclick=e=>{e.stopPropagation();showCompanyAdmin(item.id,tenant.data());};$('companies').append(manage);if(['trial','full'].includes(t.plan)&&t.ownerUid===user.uid){
    const billing=document.createElement('section');billing.className='trial-box';
    const title=document.createElement('strong');title.textContent='Abbonamento aziendale';billing.append(title);
    const select=document.createElement('select');select.setAttribute('aria-label','Piano abbonamento');
    for(const [key,p] of Object.entries(SUBSCRIPTION_PLANS)){const opt=document.createElement('option');opt.value=key;opt.textContent=p.name+' · '+formatPrice(p.priceCents)+' / 30 giorni';select.append(opt);}
    select.value=t.requestedTier||t.billing?.tier||'starter';
    const detail=document.createElement('p');const showDetails=()=>{const p=SUBSCRIPTION_PLANS[select.value],l=p.limits;detail.textContent=l.commesse+' commesse · '+l.impianti+' impianti · '+l.admin+' amministratori · '+l.operatori+' operatori · '+l.mezzi+' mezzi · '+l.squadre+' squadre · '+l.storageBytes/1000000000+' GB previsti. Attivazione previa conferma amministrativa.';};select.onchange=showDetails;showDetails();
    const upgrade=document.createElement('button');upgrade.textContent=t.upgradeRequested?'Aggiorna richiesta abbonamento':'Richiedi abbonamento';upgrade.onclick=async()=>{if(auth.currentUser?.uid!==user.uid)return;upgrade.disabled=true;try{await tenant.ref.update({upgradeRequested:true,requestedTier:select.value,upgradeRequestedAt:statics.FieldValue.serverTimestamp()});upgrade.textContent='Richiesta abbonamento inviata';}catch(err){status(userError(err));}finally{upgrade.disabled=false;}};
    billing.append(select,detail,upgrade);$('companies').append(billing);
   }}
  }
  const intent=readIntent(user.uid);
  if(intent?.type==='join'){ $('join-code').value=intent.code;$('join-panel').hidden=false;$('company-panel').hidden=true;try{const state=await requestJoin(intent.code,$('join-status'));if(request!==revision||auth.currentUser?.uid!==user.uid)return;if(state==='accepted'&&companyCount){clearIntent(user.uid);$('join-panel').hidden=true;}}catch(err){if(request!==revision)return;$('join-status').textContent=joinError(err);} }
  if(!companyCount){$('company-panel').hidden=superAdmin||intent?.type==='join';if(intent?.type==='company'&&!superAdmin){$('company-name-existing').value=intent.name;try{await createCompany(intent.name,$('company-status'));if(request===revision)await refreshAccount(auth.currentUser);return;}catch(err){$('company-status').textContent=err.code==='permission-denied'||err.code==='firestore/permission-denied'?'Account verificato, ma il server non autorizza ancora l’attivazione dell’azienda. Non registrarti di nuovo.':userError(err);}}status(superAdmin?'Apri Super Admin per amministrare le aziende.':intent?.type==='join'?'Controlla lo stato della tua richiesta.':'Crea la tua azienda oppure unisciti a un’azienda esistente.');return;}
  $('company-panel').hidden=true;if(intent?.type==='company')clearIntent(user.uid);
  status('Seleziona la tua azienda.');
 }catch(err){if(request===revision)status(userError(err));}
}
$('verify-check').onclick=async()=>{if(verificationBusy||!auth.currentUser)return;verificationBusy=true;$('verify-check').disabled=$('verify-resend').disabled=true;try{await refreshAccount(auth.currentUser);if(!$('verify-panel').hidden)$('verify-status').textContent='Email non ancora verificata. Apri il link ricevuto e riprova.';}finally{verificationBusy=false;$('verify-check').disabled=$('verify-resend').disabled=false;}};
$('verify-resend').onclick=async()=>{if(verificationBusy||!auth.currentUser)return;verificationBusy=true;$('verify-check').disabled=$('verify-resend').disabled=true;try{await auth.currentUser.sendEmailVerification();$('verify-status').textContent='Email inviata. Controlla anche la posta indesiderata.';}catch(err){$('verify-status').textContent=userError(err);}finally{verificationBusy=false;$('verify-check').disabled=$('verify-resend').disabled=false;}};
let socialBusy=false;
if(socialBridge){
 for(const group of document.querySelectorAll('[data-social-group]'))group.hidden=false;
 for(const button of document.querySelectorAll('[data-provider]'))button.onclick=async()=>{
  if(socialBusy||creatingAccount||auth.currentUser)return;
  const feedback=button.closest('[data-social-group]').querySelector('[data-social-status]');
  let intent;
  try{const mode=button.dataset.mode==='register'?registrationMode:button.dataset.mode;intent=socialIntent(mode,mode==='company'?$('company-name').value:mode==='join'?$('join-code').value:'');}
  catch(err){feedback.textContent=err.message==='COMPANY_NAME_INVALID'?'Inserisci prima il nome della tua azienda.':'Inserisci prima un codice aziendale valido.';return;}
  socialBusy=true;creatingAccount=true;const controls=Array.from(document.querySelectorAll('[data-provider],#register-submit,#login button,#join-submit,#back-to-login,#back-from-join,#join-existing'));const previous=controls.map(el=>el.disabled);controls.forEach(el=>el.disabled=true);feedback.textContent='Accesso in corso…';
  try{
   const credential=await getProviderCredential({provider:button.dataset.provider,bridge:socialBridge,firebase});
   const result=await auth.signInWithCredential(credential);
   if(intent)saveIntent(result.user.uid,intent);
   creatingAccount=false;await refreshAccount(result.user);
  }catch(err){
   feedback.textContent=err.code==='auth/account-exists-with-different-credential'?'Esiste già un account con questa email. Accedi con il metodo originale.':err.code==='auth/operation-not-allowed'?'Questo metodo di accesso non è ancora attivo.':userError(err);
   if(auth.currentUser){creatingAccount=false;await refreshAccount(auth.currentUser);}
  }finally{socialBusy=false;creatingAccount=false;controls.forEach((el,i)=>el.disabled=previous[i]);}
 };
}
auth.onAuthStateChanged(refreshAccount);
window.addEventListener('pagehide',()=>{adminPanel.clear();entry.close();},{once:true});

}
start().catch(e=>{document.getElementById('status').textContent='Avvio non disponibile. Riprova o contatta l’assistenza.';});

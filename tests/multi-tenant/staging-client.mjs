// Standalone staging client. Never imported by production index.html or app.js.
import { initializeApp } from 'https://www.gstatic.com/firebasejs/11.10.0/firebase-app.js';
import { getAuth, setPersistence, inMemoryPersistence, signInWithEmailAndPassword, onAuthStateChanged, signOut } from 'https://www.gstatic.com/firebasejs/11.10.0/firebase-auth.js';
import { initializeFirestore, memoryLocalCache, collection, doc, getDocsFromServer, getDocFromServer, query, limit } from 'https://www.gstatic.com/firebasejs/11.10.0/firebase-firestore.js';
const config=Object.freeze({apiKey:'AIzaSyC3ujdb0XyHKk-B0eGt8IVr_U_DKEXxaLI',authDomain:'varga-cantieri-v2-staging.firebaseapp.com',projectId:'varga-cantieri-v2-staging',appId:'1:394791335182:web:6c31c93ac34e64c881935f'});
const byId=id=>document.getElementById(id);
let epoch=0, busy=false, auth, db;
function status(message,error=false){byId('status').textContent=message;byId('status').dataset.error=String(error);}
function reset(){epoch++;byId('companies').replaceChildren();byId('selected').hidden=true;byId('companyName').textContent='';byId('companyRole').textContent='';}
function stillCurrent(n,uid){return n===epoch&&auth.currentUser?.uid===uid;}
function safeId(value){return typeof value==='string'&&/^[A-Za-z0-9_-]{1,128}$/.test(value)&&!/^__.*__$/.test(value);}
async function readCompany(uid,tenantId){if(!safeId(uid)||!safeId(tenantId))return null;
  const [t,m]=await Promise.all([getDocFromServer(doc(db,'tenants',tenantId)),getDocFromServer(doc(db,'tenants',tenantId,'members',uid))]);
  if(!t.exists()||!m.exists()||t.data().status!=='active'||m.data().status!=='active')return null;
  return {id:tenantId,name:String(t.data().name||tenantId),role:String(m.data().role||'')};
}
async function load(user){reset();const n=epoch;byId('login').hidden=!!user;byId('workspace').hidden=!user;
  if(!user){status('Accedi con un account di prova.');return;}
  if(!user.emailVerified){status('Email non verificata. Contatta chi gestisce lo staging.',true);return;}
  status('Verifica delle aziende autorizzate...');
  try{
    const membership=await getDocsFromServer(query(collection(db,'users',user.uid,'memberships'),limit(25)));
    const candidates=membership.docs.filter(d=>d.data().status==='active');
    const rows=await Promise.all(candidates.map(async d=>{try{return await readCompany(user.uid,d.id);}catch(e){if(e.code==='permission-denied')return null;throw e;}}));
    if(!stillCurrent(n,user.uid))return;
    for(const row of rows.filter(Boolean)){const button=document.createElement('button');button.textContent=row.name+' · '+row.role;button.addEventListener('click',async()=>{
      byId('selected').hidden=true;byId('companyName').textContent='';byId('companyRole').textContent='';
      const selection=++epoch;
      try{const current=await readCompany(user.uid,row.id);if(!stillCurrent(selection,user.uid))return;if(!current){await load(user);return;}
        byId('companyName').textContent=current.name;byId('companyRole').textContent='Ruolo: '+current.role;byId('selected').hidden=false;
      }catch{if(stillCurrent(selection,user.uid)){await load(user);status('Accesso non disponibile. Riprovare.',true);}}
    });byId('companies').append(button);}
    status(rows.some(Boolean)?'Seleziona un’azienda autorizzata.':'Nessuna azienda autorizzata.'); 
  }catch{if(stillCurrent(n,user.uid))status('Elenco non disponibile: verifica regole e connessione di staging.',true);}
}
async function start(){
  // Hard fail closed: no fallback to production or to the default Firebase app.
  if(config.projectId!=='varga-cantieri-v2-staging'||config.authDomain!=='varga-cantieri-v2-staging.firebaseapp.com'||config.appId!=='1:394791335182:web:6c31c93ac34e64c881935f')throw Error('Wrong staging config');
  const app=initializeApp(config,'varga-v2-staging');auth=getAuth(app);await setPersistence(auth,inMemoryPersistence);
  db=initializeFirestore(app,{localCache:memoryLocalCache()});
  onAuthStateChanged(auth,user=>{void load(user);},()=>{reset();byId('workspace').hidden=true;byId('login').hidden=false;status('Sessione non disponibile.',true);});
  byId('loginForm').addEventListener('submit',async e=>{e.preventDefault();if(busy)return;busy=true;byId('submit').disabled=true;status('Accesso in corso...');
    const password=byId('password').value;byId('password').value='';
    try{await signInWithEmailAndPassword(auth,byId('email').value.trim(),password);}catch{status('Accesso non riuscito. Verifica le credenziali di prova.',true);}finally{busy=false;byId('submit').disabled=false;}
  });
  byId('refresh').addEventListener('click',async()=>{const user=auth.currentUser;if(!user)return;reset();try{await user.reload();await user.getIdToken(true);if(auth.currentUser?.uid===user.uid)await load(user);}catch{status('Verifica sessione non riuscita.',true);}});
  byId('logout').addEventListener('click',async()=>{reset();byId('workspace').hidden=true;try{await signOut(auth);}catch{status('Uscita non riuscita: ricarica la pagina.',true);}});
}
start().catch(()=>{reset();status('Avvio staging non riuscito. Nessun accesso alla produzione.',true);});

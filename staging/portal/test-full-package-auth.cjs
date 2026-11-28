// Real Firebase Auth + Firestore emulators; no live accounts or data.
const assert=require('node:assert/strict'),fs=require('node:fs'),http=require('node:http'),path=require('node:path');
const {chromium}=require('playwright');
const root=path.join(__dirname,'full-runtime-dist'),projectId='varga-cantieri-v2-staging';
(async()=>{
 assert.match(process.env.FIRESTORE_EMULATOR_HOST||'',/^127\.0\.0\.1:8080$/);
 assert.match(process.env.FIREBASE_AUTH_EMULATOR_HOST||'',/^127\.0\.0\.1:9099$/);
 const {initializeApp,deleteApp}=require('firebase-admin/app');
 const {getAuth}=require('firebase-admin/auth'),{getFirestore}=require('firebase-admin/firestore');
 const withStorage=process.env.FIREBASE_STORAGE_EMULATOR_HOST==='127.0.0.1:9199';
 const admin=initializeApp({projectId,...(withStorage?{storageBucket:projectId+'.appspot.com'}:{})}),db=getFirestore(admin),auth=getAuth(admin);
 const storageHandler=withStorage?(await import('../storage-local-http.mjs')).createLocalStorageHandler((await import('../storage-quota-service.mjs')).createStorageQuotaService({app:admin})):null;
 const uid='local-admin',email='local-admin@example.test',password='Local-Test-Only-42';
 await auth.createUser({uid,email,password,emailVerified:true,displayName:'Local test admin'});
 const superUid='local-super',superEmail='local-super@example.test';
 await auth.createUser({uid:superUid,email:superEmail,password,emailVerified:true,displayName:'Local Super Admin'});
 await auth.setCustomUserClaims(superUid,{superAdmin:true});
 for(const id of ['demo-a','demo-b']){
  await db.doc('tenants/'+id).set({demo:true,status:'active',name:id,originalCreateEnabled:true,originalManualEnabled:true,originalArchivesEnabled:true,priceListEnabled:true,priceListWriteEnabled:true,snowCreateEnabled:true});
  await db.doc('tenants/'+id+'/members/'+uid).set({role:'admin',status:'active',demo:true,historicalId:'preserved-local-admin'});
  await db.doc('users/'+uid+'/memberships/'+id).set({tenantId:id,role:'admin',status:'active'});
  await db.doc('tenants/'+id+'/commesse/same-id').set({nome:id+' original',createdAt:new Date(),excelModelVersion:2});
  if(id==='demo-a'){await db.doc('tenants/'+id+'/members/'+superUid).set({role:'employee',status:'active',demo:true,historicalId:'preserved-local-super'});await db.doc('users/'+superUid+'/memberships/'+id).set({tenantId:id,role:'employee',status:'active'});}
 }
 const server=http.createServer((req,res)=>{
  if(req.url.startsWith('/__v2_documents/')){if(!storageHandler){res.writeHead(503);return res.end();}return storageHandler(req,res);}
  const rel=decodeURIComponent(new URL(req.url,'http://localhost').pathname).replace(/^\/+/,'')||'index.html',file=path.resolve(root,rel);
  if(!file.startsWith(root+path.sep)||!fs.existsSync(file)||!fs.statSync(file).isFile()){res.writeHead(404);return res.end();}
  const types={'.html':'text/html','.js':'application/javascript','.css':'text/css','.png':'image/png','.svg':'image/svg+xml','.webmanifest':'application/manifest+json'};
  res.writeHead(200,{'Content-Type':types[path.extname(file)]||'application/octet-stream'});res.end(fs.readFileSync(file));
 });
 await new Promise(resolve=>server.listen(4173,'127.0.0.1',resolve));
 const browser=await chromium.launch({headless:true});
 try{
  const context=await browser.newContext({serviceWorkers:'block'}),external=[];
  await context.route('**/*',route=>{const u=new URL(route.request().url());if(u.hostname==='127.0.0.1'&&['4173','8080','9099'].includes(u.port))return route.continue();external.push(u.hostname);return route.abort();});
  const page=await context.newPage();page.setDefaultTimeout(20000);
  const runtimeNavigations=[];
  page.on('framenavigated',frame=>{if(frame.parentFrame()!==page.mainFrame())return;try{const u=new URL(frame.url());runtimeNavigations.push({path:u.pathname,legacyReset:u.searchParams.has('cacheReset'),recovery:u.searchParams.has('firestoreRecovery')});console.log('FRAME NAVIGATION:',u.pathname,u.searchParams.has('cacheReset')?'legacy-cache-reset':u.searchParams.has('firestoreRecovery')?'recovery':'normal');}catch{}});
  // Capture only transport metadata from the synthetic localhost Firestore.
  // Never log auth headers, request bodies, passwords or token/query parameters.
  const writeTrace=[];
  const trace=(req,event,status)=>{
   const u=new URL(req.url());
   if(u.hostname!=='127.0.0.1'||u.port!=='8080'||!/(?:Write|commit)/i.test(u.pathname))return;
   writeTrace.push({event,method:req.method(),path:u.pathname,status});
   if(writeTrace.length>40)writeTrace.shift();
  };
  page.on('request',req=>trace(req,'request',null));
  page.on('response',res=>trace(res.request(),'response',res.status()));
  page.on('requestfailed',req=>trace(req,'failed',req.failure()?.errorText||'unknown'));

  await page.goto('http://127.0.0.1:4173/');
  await page.locator('#auth-options').waitFor({state:'visible'});
  assert.equal(await page.locator('#login').isHidden(),true);
  await page.locator('#open-register').click();await page.locator('#register').waitFor({state:'visible'});
  assert.equal(await page.locator('#register-trial').isVisible(),true);
  await page.locator('#back-to-login').click();
  await page.locator('#open-join').click();await page.locator('#join-code').fill('VC-AB12CD34');await page.locator('#join-submit').click();
  assert.equal(await page.locator('#register-title').textContent(),'Crea il tuo account');assert.equal(await page.locator('#company-name-field').isHidden(),true);
  await page.locator('#back-to-login').click();await page.locator('#back-from-join').click();
  // Regression: a freshly registered account stays on the verification screen,
  // then creates its company without a full reload or a second sign-in.
  const newEmail='local-registration@example.test',newCompany='Azienda nuova collaudo';
  await page.locator('#open-register').click();
  await page.locator('#company-name').fill(newCompany);
  await page.locator('#register-name').fill('Nuovo amministratore');
  await page.locator('#register-email').fill(newEmail);
  await page.locator('#register-password').fill(password);
  await page.locator('#register-confirm').fill(password);
  await page.locator('#register-submit').click();
  await page.locator('#verify-panel').waitFor({state:'visible'});
  assert.equal(await page.locator('#login-card').isHidden(),true);
  assert.equal(await page.locator('#company-panel').isHidden(),true);
  assert.equal(await page.locator('#runtime').isHidden(),true);
  await page.locator('#verify-check').click();
  await page.locator('#verify-status').filter({hasText:'non ancora verificata'}).waitFor();
  const fresh=await auth.getUserByEmail(newEmail);
  assert.equal(fresh.emailVerified,false);
  await page.evaluate(()=>{window.__registrationPageMarker='same-session';});
  // A conflicting code denies the entire bootstrap batch. Retrying must reuse
  // the same persisted company ID; no partial company or duplicate.
  const activation={type:'company',name:newCompany,tenantId:'demo-9e43f616-410f-4af4-9ed1-5b83502a7fc1',joinCode:'VC-FEEDABCD'};
  await db.doc('tenantJoinCodes/'+activation.joinCode).set({tenantId:'demo-a',companyName:'demo-a',status:'active'});
  await page.evaluate(({uid,intent})=>localStorage.setItem('varga.onboarding.'+uid,JSON.stringify(intent)),{uid:fresh.uid,intent:activation});
  await auth.updateUser(fresh.uid,{emailVerified:true});
  await page.locator('#verify-check').click();
  await page.locator('#company-status').filter({hasText:'non autorizza ancora'}).waitFor();
  assert.equal((await db.doc('tenants/'+activation.tenantId).get()).exists,false);
  assert.deepEqual(await page.evaluate(uid=>JSON.parse(localStorage.getItem('varga.onboarding.'+uid)),fresh.uid),activation);
  await db.doc('tenantJoinCodes/'+activation.joinCode).delete();
  await page.locator('#company-form button[type=submit]').click();
  await page.locator('[data-tenant="'+activation.tenantId+'"]').waitFor();
  assert.equal(await page.evaluate(()=>window.__registrationPageMarker),'same-session');
  const created=await db.collection('tenants').where('ownerUid','==',fresh.uid).get();
  assert.equal(created.size,1);
  const trial=created.docs[0].data();
  assert.deepEqual(trial.limits,{commesse:5,impianti:100,admin:1,operatori:11,mezzi:11,squadre:2,storageBytes:100000000,personale:11});
  assert.equal(trial.trialDays,30);
  assert.ok(trial.trialStartedAt.toMillis()<=Date.now());
  await page.getByRole('combobox',{name:'Piano abbonamento',exact:true}).selectOption('professional');
  await page.getByRole('button',{name:'Richiedi abbonamento',exact:true}).click();
  await page.getByRole('button',{name:'Richiesta abbonamento inviata',exact:true}).waitFor();
  assert.equal((await created.docs[0].ref.get()).data().requestedTier,'professional');
  assert.equal((await created.docs[0].ref.get()).data().upgradeRequested,true);
  // Exercise actual tenant-scoped client writes with Auth/Firestore emulators.
  await page.locator('[data-tenant="'+activation.tenantId+'"]').click();
  await page.waitForFunction(()=>{try{return window.__vargaLocalRuntime().session.session.context.tenantId.startsWith('demo-9e43');}catch{return false;}});
  await page.frameLocator('#runtime').locator('#side-menu').waitFor({state:'attached'});
  const trialFrame=await (await page.locator('#runtime').elementHandle()).contentFrame();
  await trialFrame.waitForURL('**/original.html');
  await trialFrame.waitForFunction(()=>document.readyState==='complete'&&Boolean(window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__?.isActive()));
  const createTrialJob=async()=>trialFrame.evaluate(async()=>{const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__,stamp=window.firebase.firestore.FieldValue.serverTimestamp();const ref=await s.session.firestore.collection('commesse').add({nome:'Quota test',codice:'',excelModelVersion:2,priceListVersion:2,workItemsModelVersion:2,percentualeRibassoGenerale:0.01,excelModelActivatedAt:stamp,excelModelActivatedBy:s.auth.currentUser.email,nextImpiantoNumber:1,parentCommessaId:null,creatoDa:s.auth.currentUser.email,createdAt:stamp});return ref.id;});
  if(withStorage){
   await trialFrame.locator('#open-private-docs-btn').evaluate(b=>b.click());
   const docs=page.locator('#company-documents');await docs.waitFor({state:'visible'});
   await docs.getByText('Nessun documento disponibile.').waitFor();
   const bytes=Buffer.from('%PDF-1.4\nSynthetic UI test\n%%EOF\n');
   await docs.locator('input[type=file]').setInputFiles({name:'prova.pdf',mimeType:'application/pdf',buffer:bytes});
   await docs.getByRole('button',{name:'Carica documento',exact:true}).click();
   await docs.getByText(/prova\.pdf/).waitFor();
   assert.equal((await created.docs[0].ref.get()).data().usage.storageBytes,bytes.length);
   await docs.getByText('prova.pdf',{exact:false}).filter({hasText:'Privato'}).waitFor();
   page.once('dialog',d=>d.accept());
   await docs.getByRole('button',{name:'Condividi con azienda',exact:true}).click();
   await docs.getByRole('button',{name:'Rendi privato',exact:true}).waitFor();
   const sharedDoc=(await created.docs[0].ref.collection('storedFiles').where('name','==','prova.pdf').get()).docs[0];
   assert.equal(sharedDoc.data().visibility,'company');
   await docs.getByRole('button',{name:'Rendi privato',exact:true}).click();
   await docs.getByRole('button',{name:'Condividi con azienda',exact:true}).waitFor();
   assert.equal((await sharedDoc.ref.get()).data().visibility,'private');
   page.once('dialog',d=>d.accept());
   await docs.getByRole('button',{name:'Elimina',exact:true}).click();
   await docs.getByText('Nessun documento disponibile.').waitFor();
   assert.equal((await created.docs[0].ref.get()).data().usage.storageBytes,0);
   await docs.getByRole('button',{name:'Torna al cantiere',exact:true}).click();
   await docs.waitFor({state:'hidden'});
   console.log('DOCUMENT UI PASS: original menu, private upload, confirmed deletion and quota recovery.');
  }
  // First-user operational setup through the real UI: personnel, vehicle and commessa.
  const openMenu=async()=>{const menu=trialFrame.locator('#side-menu');if(!(await menu.isVisible())){await trialFrame.locator('#menu-toggle-btn').click();await menu.waitFor({state:'visible'});}};
  await openMenu();
  await trialFrame.locator('#open-panel-personale').click();
  await trialFrame.locator('#personnel-v2').waitFor({state:'visible'});
  await trialFrame.locator('#personnel-v2 [data-new="personale"]').click();
  await trialFrame.locator('#registry-editor').waitFor({state:'visible'});
  await trialFrame.locator('#registry-editor input[name="nome"]').fill('Operatore');
  await trialFrame.locator('#registry-editor input[name="cognome"]').fill('Collaudo');
  await trialFrame.locator('#registry-editor input[name="mansione"]').fill('Operatore cantiere');
  await trialFrame.locator('#registry-editor').locator('button.btn-primary').click();
  await trialFrame.waitForFunction(async()=>{const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__;return (await s.session.firestore.collection('personale').where('nome','==','Operatore').where('cognome','==','Collaudo').get()).size===1;});
  console.log('FIRST USER PERSONALE PASS: Operatore Collaudo saved from the Personale v2 form.');
  await trialFrame.locator('#management-close-btn').click();
  await trialFrame.locator('#management-page').waitFor({state:'hidden'});
  await openMenu();
  await trialFrame.locator('#open-panel-mezzi').click();
  console.log('FIRST USER MEZZI UI STATE '+JSON.stringify(await trialFrame.evaluate(()=>({menu:document.getElementById('side-menu')?.className,management:document.getElementById('management-page')?.className,managementAria:document.getElementById('management-page')?.getAttribute('aria-hidden'),panel:document.getElementById('panel-mezzi')?.className,vehicles:document.getElementById('vehicles-v2')?.className,admin:typeof canManageData==='function'?canManageData():null}))));
  await trialFrame.locator('#vehicles-v2').waitFor({state:'visible'});
  await trialFrame.locator('#vehicles-v2 [data-new="mezzi"]').click();
  await trialFrame.locator('#registry-editor').waitFor({state:'visible'});
  await trialFrame.locator('#registry-editor input[name="codiceMezzo"]').fill('MEZZO-COLLAUDO-01');
  await trialFrame.locator('#registry-editor input[name="marca"]').fill('Varga');
  await trialFrame.locator('#registry-editor input[name="modello"]').fill('Furgone');
  await trialFrame.locator('#registry-editor').locator('button.btn-primary').click();
  await trialFrame.waitForFunction(async()=>{const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__;return (await s.session.firestore.collection('mezzi').where('codiceMezzo','==','MEZZO-COLLAUDO-01').get()).size===1;});
  console.log('FIRST USER MEZZO PASS: MEZZO-COLLAUDO-01 saved from the Mezzi v2 form.');
  await trialFrame.locator('#management-close-btn').click();
  await trialFrame.locator('#management-page').waitFor({state:'hidden'});
  await openMenu();
  await trialFrame.locator('#open-panel-commesse').click();
  await trialFrame.locator('#open-new-commessa-btn').click();
  await trialFrame.locator('#commessa-form').waitFor({state:'visible'});
  await trialFrame.locator('#commessa-name').fill('Cantiere Collaudo Prima Configurazione');
  await trialFrame.locator('#commessa-code').fill('COL-001');
  await trialFrame.locator('#commessa-form').locator('button[type="submit"]').click();
  await trialFrame.waitForFunction(async()=>{const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__;return (await s.session.firestore.collection('commesse').where('nome','==','Cantiere Collaudo Prima Configurazione').get()).size===1;});
  const uiJob=(await created.docs[0].ref.collection('commesse').where('nome','==','Cantiere Collaudo Prima Configurazione').get()).docs[0].id;
  console.log('FIRST USER COMMESSA PASS: Cantiere Collaudo Prima Configurazione saved from the Gestione commesse form.');
  let firstJob=uiJob;
  for(let i=0;i<4;i++){const id=await createTrialJob();if(i===0&& !firstJob)firstJob=id;}
  await assert.rejects(createTrialJob);
  assert.equal((await created.docs[0].ref.collection('commesse').get()).size,5);
  assert.equal((await created.docs[0].ref.get()).data().usage.commesse,5);
  // Seed 99 synthetic plant documents and 10 vehicles to exercise the exact boundary.
  const seeded=db.batch();
  for(let i=0;i<99;i++)seeded.set(created.docs[0].ref.collection('commesse').doc(firstJob).collection('impianti').doc('seed-'+i),{denominazione:'Synthetic '+i});
  for(let i=0;i<10;i++)seeded.set(created.docs[0].ref.collection('mezzi').doc('seed-'+i),{nome:'Synthetic vehicle '+i});
  seeded.update(created.docs[0].ref,{'usage.impianti':99,'usage.mezzi':10});
  await seeded.commit();
  if(withStorage){
   // The candidate plant-card action calls this exact company-scoped bridge.
   assert.ok(fs.readFileSync(path.join(root,'app.js'),'utf8').includes('photoButton.dataset.v2PlantPhotos'));
   assert.equal(await trialFrame.evaluate(jobId=>window.parent.__vargaOpenCompanyDocuments({archive:'commesse',jobId,plantId:'seed-0'}),firstJob),true);
   const docs=page.locator('#company-documents');await docs.getByRole('heading',{name:'Fotografie impianto'}).waitFor();
   await docs.getByText('Nessun documento disponibile.').waitFor();
   const png=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a0i0AAAAASUVORK5CYII=','base64');
   await docs.locator('input[type=file]').setInputFiles({name:'impianto.png',mimeType:'image/png',buffer:png});
   await docs.getByRole('combobox',{name:'Visibilità documento'}).selectOption('company');
   await docs.getByRole('button',{name:'Carica documento',exact:true}).click();
   await docs.getByText(/impianto\.png/).waitFor();
   const saved=(await created.docs[0].ref.collection('storedFiles').where('name','==','impianto.png').get()).docs[0].data();
   assert.deepEqual(saved.plant,{archive:'commesse',jobId:firstJob,plantId:'seed-0'});assert.equal(saved.visibility,'company');
   assert.equal((await created.docs[0].ref.get()).data().usage.storageBytes,png.length);
   page.once('dialog',d=>d.accept());await docs.getByRole('button',{name:'Elimina',exact:true}).click();
   await docs.getByText('Nessun documento disponibile.').waitFor();
   assert.equal((await created.docs[0].ref.get()).data().usage.storageBytes,0);
   await docs.getByRole('button',{name:'Torna al cantiere',exact:true}).click();
   console.log('PLANT PHOTO UI PASS: verified company bridge, scoped image upload, shared visibility and quota release.');
  }
  const plant=()=>trialFrame.evaluate(async job=>{const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__,stamp=window.firebase.firestore.FieldValue.serverTimestamp(),ref=s.session.firestore.collection('commesse').doc(job).collection('impianti').doc();await ref.set({id:ref.id,commessaId:job,nome:'Quota plant',createdAt:stamp,createdBy:s.auth.currentUser.uid});},firstJob);
  await plant();await assert.rejects(plant);
  const vehicle=()=>trialFrame.evaluate(()=>window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__.session.firestore.collection('mezzi').add({nome:'Quota vehicle'}).then(()=>true));
  await vehicle();await assert.rejects(vehicle);
  assert.equal((await created.docs[0].ref.get()).data().usage.impianti,100);
  assert.equal((await created.docs[0].ref.get()).data().usage.mezzi,11);
  // Actual original squad payload: one current document per job with multiple rows,
  // plus date-keyed history. The quota counts rows across all current documents.
  const secondJob=(await created.docs[0].ref.collection('commesse').get()).docs.find(d=>d.id!==firstJob).id;
  const squad=(job,date,count)=>trialFrame.evaluate(async({job,date,count})=>{
   const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__,f=s.session.firestore;
   const current=f.collection('squadreCommesse').doc(job),history=f.collection('squadreStorico').doc(date+'__'+job);
   const payload={commessaId:job,commessaNome:'Quota test',riferimentoData:date,squadre:Array.from({length:count},(_,i)=>({caposquadra:'Synthetic lead '+i,personale:'Synthetic operator '+i})),updatedAt:window.firebase.firestore.FieldValue.serverTimestamp(),updatedBy:s.auth.currentUser.email};
   await f.runTransaction(async tx=>{const h=await tx.get(history),next={...payload,existingSquadreCountBeforeSave:h.exists?h.data().squadre.length:0};tx.set(current,next,{merge:true});tx.set(history,{...next,dateKey:date},{merge:true});});
  },{job,date,count});
  await squad(firstJob,'2026-11-28',2);
  assert.equal((await created.docs[0].ref.get()).data().usage.squadre,2);
  await assert.rejects(()=>squad(secondJob,'2026-11-28',1));
  assert.equal((await created.docs[0].ref.collection('squadreCommesse').doc(secondJob).get()).exists,false);
  await squad(firstJob,'2026-11-29',1);
  await squad(secondJob,'2026-11-29',1);
  assert.equal((await created.docs[0].ref.get()).data().usage.squadre,2);
  assert.equal((await created.docs[0].ref.collection('squadreStorico').get()).size,3);
  assert.equal((await created.docs[0].ref.collection('squadreStorico').doc('2026-11-28__'+firstJob).get()).data().squadre.length,2);
  await assert.rejects(()=>trialFrame.evaluate(async job=>{
   const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__,f=s.session.firestore;
   await f.collection('squadreCommesse').doc(job).update({squadre:[{caposquadra:'Uno'},{caposquadra:'Due'},{caposquadra:'Tre'}],updatedAt:window.firebase.firestore.FieldValue.serverTimestamp()});
  },firstJob));
  assert.equal((await created.docs[0].ref.get()).data().usage.squadre,2);
  console.log('SQUAD QUOTA: actual rows across jobs, historical days excluded, replacement releases one slot, unpaired update denied.');
  // Expiry denies writes while read access remains intact.
  await created.docs[0].ref.update({trialStartedAt:new Date(Date.now()-31*86400000)});
  await assert.rejects(()=>trialFrame.evaluate(()=>window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__.session.firestore.collection('mezzi').doc('seed-0').update({nome:'Denied after expiry'})));
  assert.equal(await trialFrame.evaluate(async()=>{const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__;return (await s.session.firestore.collection('mezzi').doc('seed-0').get({source:'server'})).exists;}),true);
  await page.locator('#company-return').click();
  console.log('TRIAL RESOURCE QUOTAS: 5 jobs, 100 plants, 11 vehicles, expired update denied, reads retained.');

  assert.equal((await db.doc('tenants/'+created.docs[0].id+'/members/'+fresh.uid).get()).data().role,'admin');
  assert.equal((await db.doc('users/'+fresh.uid+'/memberships/'+created.docs[0].id).get()).data().status,'active');
  await page.locator('#logout').click();await page.locator('#auth-options').waitFor({state:'visible'});
  await page.locator('#open-login').click();
  await page.locator('#email').fill(newEmail);await page.locator('#password').fill(password);await page.locator('#login button[type=submit]').click();
  await page.locator('[data-tenant]').filter({hasText:newCompany}).waitFor();
  assert.equal((await db.collection('tenants').where('ownerUid','==',fresh.uid).get()).size,1);
  await page.locator('#logout').click();await page.locator('#auth-options').waitFor({state:'visible'});
  console.log('REGISTRATION SESSION: unverified gate, verified activation, same page, owner membership and repeat login.');
  await page.locator('#open-login').click();
  await page.locator('#email').fill(email);await page.locator('#password').fill(password);await page.locator('#login button[type=submit]').click();
  await page.locator('[data-tenant="demo-a"]').waitFor({state:'visible'});
  assert.equal(await page.locator('[data-tenant="demo-b"]').count(),1,'Verified company list must be published atomically before selection');
  await page.locator('[data-tenant="demo-a"]').click();
  const frame=page.frameLocator('#runtime');
  await frame.locator('#side-menu').waitFor({state:'attached'});
  await page.waitForFunction(()=>{try{return window.__vargaLocalRuntime().session.session.context.tenantId==='demo-a';}catch{return false;}});
  const a=await page.evaluate(async()=>{const s=window.__vargaLocalRuntime().session;window.__oldLocalScope=s;return (await s.session.firestore.collection('commesse').doc('same-id').get({source:'server'})).data().nome;});
  assert.equal(a,'demo-a original');
  // Writes from the actual iframe realm, using the selected company Firestore facade.
  const child=await (await page.locator('#runtime').elementHandle()).contentFrame();assert(child);await child.waitForFunction(()=>Boolean(window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__?.isActive()));
  console.log('REAL AUTH AND ORIGINAL FRAME: verified demo-a membership and isolated document read; frame URL '+new URL(child.url()).pathname);
  try{
  await child.evaluate(async()=>{const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__;const stamp=window.firebase.firestore.FieldValue.serverTimestamp();await s.session.firestore.collection('commesse').add({nome:'Created from actual frame',codice:'',excelModelVersion:2,priceListVersion:2,workItemsModelVersion:2,percentualeRibassoGenerale:0.01,excelModelActivatedAt:stamp,excelModelActivatedBy:s.auth.currentUser.email,nextImpiantoNumber:1,parentCommessaId:null,creatoDa:s.auth.currentUser.email,createdAt:stamp});});
  }catch(error){
   // A rejected create can still have reached the emulator. Diagnose, never retry
   // the write and never turn a failed operation into an apparent success.
   try{
    const records=await db.collection('tenants/demo-a/commesse').get();
    const matching=records.docs.filter(d=>d.data().nome==='Created from actual frame');
    console.error('SYNTHETIC WRITE DIAGNOSTIC '+JSON.stringify({errorCode:error.code||null,total:records.size,matchingCreated:matching.length,otherCompanyCount:(await db.collection('tenants/demo-b/commesse').get()).size,writeTrace}));
   }catch(diagnosticError){console.error('SYNTHETIC WRITE DIAGNOSTIC unavailable:',diagnosticError.code||'unknown');}
   throw error;
  }
  assert.equal((await db.collection('tenants/demo-a/commesse').get()).size,2);
  console.log('REAL FRAME WRITE: company A persisted the original commessa payload through scoped Firestore.');
  assert.equal(await page.locator('body').evaluate(el=>el.classList.contains('company-runtime')),true);
  await page.locator('#company-return').click();
  await page.locator('[data-tenant="demo-b"]').click();
  await page.waitForFunction(()=>{try{return window.__vargaLocalRuntime().session.session.context.tenantId==='demo-b';}catch{return false;}});
  assert.equal(await page.evaluate(async()=>{const s=window.__vargaLocalRuntime().session;return (await s.session.firestore.collection('commesse').doc('same-id').get({source:'server'})).data().nome;}),'demo-b original');
  assert.equal(await page.evaluate(()=>{try{window.__oldLocalScope.session.firestore.collection('commesse');return false;}catch{return true;}}),true);
  assert.equal((await db.collection('tenants/demo-b/commesse').get()).size,1);
  await page.locator('#company-return').click();await page.locator('#logout').click();await page.locator('#auth-options').waitFor({state:'visible'});
  assert.equal(await page.locator('#runtime').isHidden(),true);
  assert.equal(await page.locator('#admin-panel').isHidden(),true);
  await page.locator('#open-login').click();await page.locator('#email').fill(superEmail);await page.locator('#password').fill(password);await page.locator('#login button[type=submit]').click();
  await page.locator('[data-tenant="demo-a"]').waitFor({state:'visible'});
  assert.equal(await page.locator('#admin-panel').isHidden(),true);
  assert.equal(await page.locator('[data-tenant="demo-b"]').count(),0);
  await page.locator('[data-tenant="demo-a"]').click();
  await page.waitForFunction(()=>{try{return window.__vargaLocalRuntime().session.session.context.tenantId==='demo-a';}catch{return false;}});
  assert.equal(await page.evaluate(()=>window.__vargaLocalRuntime().session.session.context.role),'employee');
  assert.equal(await page.locator('body').evaluate(el=>el.classList.contains('company-runtime')),true);
  await page.locator('#company-return').click();
  await page.locator('#admin-toggle').click();
  const arow=page.locator('[data-admin-tenant="demo-a"]');
  await arow.waitFor({state:'visible'});
  await arow.getByRole('button',{name:'Utenti',exact:true}).click();
  const member=page.locator('[data-member="local-admin"]');
  await member.getByRole('combobox').selectOption('supervisor');
  await member.getByRole('button',{name:'Salva ruolo'}).click();
  await page.waitForFunction(()=>document.getElementById('admin-status')?.textContent==='Modifica salvata.');
  assert.equal((await db.doc('tenants/demo-a/members/'+uid).get()).data().role,'supervisor');
  assert.equal((await db.doc('users/'+uid+'/memberships/demo-a').get()).data().role,'supervisor');
  assert.equal((await db.doc('tenants/demo-b/members/'+uid).get()).data().role,'admin');
  assert.equal((await db.doc('tenants/demo-a/members/'+uid).get()).data().historicalId,'preserved-local-admin');
  const trialRow=page.locator('[data-admin-tenant="'+activation.tenantId+'"]');
  await trialRow.getByText(/Prova scadut|Piano prova scaduto/).waitFor();
  await trialRow.getByRole('combobox',{name:'Abbonamento da attivare',exact:true}).selectOption('professional');
  await trialRow.getByRole('textbox',{name:'Riferimento pagamento',exact:true}).fill('TEST-PAYMENT-0001');
  await trialRow.getByRole('button',{name:'Registra pagamento e attiva',exact:true}).click();
  await page.locator('#admin-status').filter({hasText:'Verifica l’incasso'}).waitFor();
  assert.equal((await created.docs[0].ref.get()).data().plan,'trial');
  await trialRow.getByRole('checkbox').check();
  await trialRow.getByRole('button',{name:'Registra pagamento e attiva',exact:true}).click();
  await trialRow.getByRole('button',{name:'Registra rinnovo',exact:true}).waitFor();
  const upgraded=(await created.docs[0].ref.get()).data();
  assert.equal(upgraded.plan,'full');assert.equal(upgraded.limits.commesse,50);assert.equal(upgraded.usage.commesse,5);
  assert.equal(upgraded.billing.amountCents,3990);assert.equal(upgraded.billing.tier,'professional');assert.equal(upgraded.billing.periods,1);
  assert.equal(upgraded.upgradeRequested,false);assert.equal(upgraded.planUpdatedBy,superUid);
  assert.equal((await created.docs[0].ref.collection('billingReceipts').get()).size,1);
  await trialRow.getByRole('textbox',{name:'Riferimento pagamento',exact:true}).fill('TEST-PAYMENT-0001');
  await trialRow.getByRole('checkbox').check();
  await trialRow.getByRole('button',{name:'Registra rinnovo',exact:true}).click();
  await page.locator('#admin-status').filter({hasText:'Operazione non riuscita'}).waitFor();
  assert.equal((await created.docs[0].ref.get()).data().billing.periods,1);
  await trialRow.getByRole('textbox',{name:'Riferimento pagamento',exact:true}).fill('TEST-PAYMENT-0002');
  await trialRow.getByRole('button',{name:'Registra rinnovo',exact:true}).click();
  await page.locator('#admin-status').filter({hasText:'Modifica salvata.'}).waitFor();
  const renewed=(await created.docs[0].ref.get()).data();
  assert.equal(renewed.billing.periods,2);assert.equal(renewed.billing.periodStart.toMillis(),upgraded.billing.periodStart.toMillis());
  assert.equal((await created.docs[0].ref.collection('billingReceipts').get()).size,2);
  console.log('COMMERCIAL BILLING: manual confirmation required, professional activated, duplicate receipt denied, renewal preserves paid period.');
  // Synthetic operator memberships; Professional permits three admins total.
  for(const id of ['local-paid-role-1','local-paid-role-2','local-paid-role-3']){
   await created.docs[0].ref.collection('members').doc(id).set({uid:id,role:'employee',status:'active',demo:true});
   await db.doc('users/'+id+'/memberships/'+activation.tenantId).set({tenantId:activation.tenantId,role:'employee',status:'active',demo:true});
  }
  await created.docs[0].ref.update({'usage.operatori':3});
  await trialRow.getByRole('button',{name:'Utenti',exact:true}).click();
  for(const id of ['local-paid-role-1','local-paid-role-2']){
   const m=page.locator('[data-member="'+id+'"]');
   await m.getByRole('combobox').selectOption('admin');
   await m.getByRole('button',{name:'Salva ruolo'}).click();
   await page.waitForFunction(()=>document.getElementById('admin-status')?.textContent==='Modifica salvata.');
   assert.equal((await created.docs[0].ref.collection('members').doc(id).get()).data().role,'admin');
   assert.equal((await db.doc('users/'+id+'/memberships/'+activation.tenantId).get()).data().role,'admin');
  }
  const extra=page.locator('[data-member="local-paid-role-3"]');
  await extra.getByRole('combobox').selectOption('admin');
  await extra.getByRole('button',{name:'Salva ruolo'}).click();
  await page.locator('#admin-status').filter({hasText:'Operazione non riuscita'}).waitFor();
  assert.equal((await created.docs[0].ref.collection('members').doc('local-paid-role-3').get()).data().role,'employee');
  assert.equal((await created.docs[0].ref.get()).data().usage.admin,3);
  assert.equal((await created.docs[0].ref.get()).data().usage.operatori,1);
  console.log('PAID ADMIN QUOTA: Professional allows three admins; fourth denied, mirror and counters consistent.');

  await page.locator('#admin-new-company-name').fill('Nuova azienda collaudo');
  await page.getByRole('button',{name:'Nuova azienda',exact:true}).click();
  await page.locator('[data-admin-tenant]').filter({hasText:'Nuova azienda collaudo'}).waitFor();
  assert.equal((await db.collection('tenants').where('name','==','Nuova azienda collaudo').get()).size,1);
  await page.locator('#logout').click();await page.locator('#auth-options').waitFor({state:'visible'});
  assert.equal(await page.locator('#admin-panel').isHidden(),true);
  // New employee uses the company code without being able to read the tenant before approval.
  const joinUid='local-join',joinEmail='local-join@example.test',code='VC-AB12CD34';
  await auth.createUser({uid:joinUid,email:joinEmail,password,emailVerified:true,displayName:'Local join employee'});
  await db.doc('tenantJoinCodes/'+code).set({tenantId:'demo-b',companyName:'demo-b',status:'active'});
  await db.doc('tenants/demo-b').update({joinCode:code,plan:'trial',trialDays:30,trialStartedAt:new Date(),limits:{commesse:5,impianti:100,admin:1,operatori:11,mezzi:11,squadre:2,storageBytes:100000000,personale:11},usage:{commesse:1,impianti:0,admin:1,operatori:10,mezzi:0,squadre:0,storageBytes:0,personale:0}});
  for(const reserved of Array.from({length:10},(_,i)=>'local-reserved-'+i)){
   await db.doc('tenants/demo-b/members/'+reserved).set({uid:reserved,role:'employee',status:'active',demo:true});
   await db.doc('users/'+reserved+'/memberships/demo-b').set({tenantId:'demo-b',role:'employee',status:'active',demo:true});
  }
  await page.locator('#open-join').click();await page.locator('#join-code').fill(code);await page.locator('#join-existing').click();
  await page.locator('#email').fill(joinEmail);await page.locator('#password').fill(password);await page.locator('#login button[type=submit]').click();
  await page.waitForFunction(()=>document.getElementById('join-status')?.textContent.includes('Richiesta inviata'));
  assert.equal((await db.doc('tenants/demo-b/joinRequests/'+joinUid).get()).data().status,'pending');
  assert.equal(await page.locator('[data-tenant="demo-b"]').count(),0);
  // Keep the operator signed in while a second browser tab approves the request.
  const approvingPage=await context.newPage();
  try{
   await approvingPage.goto('http://127.0.0.1:4173/');
   await approvingPage.locator('#open-login').click();await approvingPage.locator('#email').fill(email);await approvingPage.locator('#password').fill(password);await approvingPage.locator('#login button[type=submit]').click();
   await approvingPage.locator('[data-tenant="demo-b"] + button').click();
   await approvingPage.getByRole('button',{name:'Accetta',exact:true}).click();
   await approvingPage.waitForFunction(()=>document.getElementById('company-admin-status')?.textContent.includes('nessuna richiesta'));
   assert.equal((await db.doc('tenants/demo-b/members/'+joinUid).get()).data().role,'employee');
   assert.equal((await db.doc('tenants/demo-b').get()).data().usage.operatori,11);
   const fourth='local-twelfth-operator';
   await db.doc('tenants/demo-b/joinRequests/'+fourth).set({uid:fourth,email:'local-fourth@example.test',displayName:'Dodicesimo operatore',status:'pending',joinCode:code,requestedAt:new Date()});
   await approvingPage.locator('[data-tenant="demo-b"] + button').click();
   await approvingPage.getByRole('button',{name:'Accetta',exact:true}).click();
   await approvingPage.locator('#company-admin-status').filter({hasText:'Accettazione non consentita'}).waitFor();
   assert.equal((await db.doc('tenants/demo-b/members/'+fourth).get()).exists,false);
   assert.equal((await db.doc('users/'+fourth+'/memberships/demo-b').get()).exists,false);
   assert.equal((await db.doc('tenants/demo-b/joinRequests/'+fourth).get()).data().status,'pending');
   assert.equal((await db.doc('tenants/demo-b').get()).data().usage.operatori,11);
   console.log('TRIAL OPERATOR QUOTA: eleventh accepted atomically, twelfth denied, all records unchanged.');
  }finally{await approvingPage.close();}
  await page.locator('#back').click();
  await page.locator('[data-tenant="demo-b"]').waitFor({state:'visible'});
  await page.locator('#join-panel').waitFor({state:'hidden'});
  assert.equal(await page.evaluate(uid=>localStorage.getItem('varga.onboarding.'+uid),joinUid),null);
  console.log('JOIN APPROVAL: operator refreshed membership in the same authenticated session.');
  await page.locator('#logout').click();await page.locator('#auth-options').waitFor({state:'visible'});
  await page.locator('#open-login').click();await page.locator('#email').fill(joinEmail);await page.locator('#password').fill(password);await page.locator('#login button[type=submit]').click();
  await page.locator('[data-tenant="demo-b"]').waitFor({state:'visible'});
  await page.locator('#logout').click();await page.locator('#auth-options').waitFor({state:'visible'});
  await page.locator('#open-login').click();await page.locator('#email').fill(newEmail);await page.locator('#password').fill(password);await page.locator('#login button[type=submit]').click();
  await page.locator('[data-tenant="'+activation.tenantId+'"]').click();
  await page.waitForFunction(()=>{try{return window.__vargaLocalRuntime().session.session.context.tenantId.startsWith('demo-9e43');}catch{return false;}});
  await page.frameLocator('#runtime').locator('#side-menu').waitFor({state:'attached'});
  const paidFrame=await (await page.locator('#runtime').elementHandle()).contentFrame();
  await paidFrame.waitForURL('**/original.html');
  await paidFrame.waitForFunction(()=>document.readyState==='complete'&&Boolean(window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__?.isActive()));
  const updatePaid=()=>paidFrame.evaluate(()=>window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__.session.firestore.collection('mezzi').doc('seed-0').update({nome:'Paid vehicle'}));
  await updatePaid();
  // Only test administrator simulates passage of time, never actual client.
  await created.docs[0].ref.update({'billing.periodStart':new Date(Date.now()-61*86400000)});
  await assert.rejects(updatePaid);
  assert.equal(await paidFrame.evaluate(async()=>{const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__;return (await s.session.firestore.collection('mezzi').doc('seed-0').get({source:'server'})).exists;}),true);
  await page.locator('#company-return').click();await page.locator('#back').click();
  await page.locator('[data-tenant="'+activation.tenantId+'"]').filter({hasText:'scaduto'}).waitFor();
  await page.locator('#logout').click();await page.locator('#auth-options').waitFor({state:'visible'});
  console.log('PAID EXPIRY: valid period permits writes; expired renewal denied, data readable.');
  assert.ok(runtimeNavigations.some(n=>n.path==='/original.html'),'Company runtime actually navigated');
  assert.equal(runtimeNavigations.some(n=>n.legacyReset),false,'Legacy standalone PWA cache reset must not navigate embedded runtime');
  console.log('V2 FRAME REGRESSION: company runtime loaded without legacy cache-reset navigation.');
  assert.equal(external.length,0,'Unexpected external requests: '+external.join(','));
  console.log('INTEGRATED SUPER ADMIN UI: real Auth claim, tenant creation, mirrored role change, other company unchanged, historical ID preserved, logout.');
  console.log('FULL PACKAGE REAL AUTH E2E: login, verified membership, original UI, frame write, tenant switch, old scope closed, logout.');
 }finally{await browser.close();await new Promise(resolve=>server.close(resolve));await deleteApp(admin);}
})().catch(e=>{console.error(e);process.exitCode=1;});

// Real Firebase Auth + Firestore emulators; no live accounts or data.
const assert=require('node:assert/strict'),fs=require('node:fs'),http=require('node:http'),path=require('node:path');
const {chromium}=require('playwright');
const root=path.join(__dirname,'full-runtime-dist'),projectId='varga-cantieri-v2-staging';
(async()=>{
 assert.match(process.env.FIRESTORE_EMULATOR_HOST||'',/^127\.0\.0\.1:8080$/);
 assert.match(process.env.FIREBASE_AUTH_EMULATOR_HOST||'',/^127\.0\.0\.1:9099$/);
 const {initializeApp,deleteApp}=require('firebase-admin/app');
 const {getAuth}=require('firebase-admin/auth'),{getFirestore}=require('firebase-admin/firestore');
 const admin=initializeApp({projectId}),db=getFirestore(admin),auth=getAuth(admin);
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
 }
 const server=http.createServer((req,res)=>{
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
  await page.locator('#email').fill(email);await page.locator('#password').fill(password);await page.locator('#login button').click();
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
  await page.locator('[data-tenant="demo-b"]').click();
  await page.waitForFunction(()=>{try{return window.__vargaLocalRuntime().session.session.context.tenantId==='demo-b';}catch{return false;}});
  assert.equal(await page.evaluate(async()=>{const s=window.__vargaLocalRuntime().session;return (await s.session.firestore.collection('commesse').doc('same-id').get({source:'server'})).data().nome;}),'demo-b original');
  assert.equal(await page.evaluate(()=>{try{window.__oldLocalScope.session.firestore.collection('commesse');return false;}catch{return true;}}),true);
  assert.equal((await db.collection('tenants/demo-b/commesse').get()).size,1);
  await page.locator('#logout').click();await page.locator('#login').waitFor({state:'visible'});
  assert.equal(await page.locator('#runtime').isHidden(),true);
  assert.equal(await page.locator('#admin-panel').isHidden(),true);
  await page.locator('#email').fill(superEmail);await page.locator('#password').fill(password);await page.locator('#login button').click();
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
  await page.locator('#admin-new-company-name').fill('Nuova azienda collaudo');
  await page.getByRole('button',{name:'Nuova azienda',exact:true}).click();
  await page.locator('[data-admin-tenant]').filter({hasText:'Nuova azienda collaudo'}).waitFor();
  assert.equal((await db.collection('tenants').where('name','==','Nuova azienda collaudo').get()).size,1);
  await page.locator('#logout').click();await page.locator('#login').waitFor({state:'visible'});
  assert.equal(await page.locator('#admin-panel').isHidden(),true);
  assert.equal(external.length,0,'Unexpected external requests: '+external.join(','));
  console.log('INTEGRATED SUPER ADMIN UI: real Auth claim, tenant creation, mirrored role change, other company unchanged, historical ID preserved, logout.');
  console.log('FULL PACKAGE REAL AUTH E2E: login, verified membership, original UI, frame write, tenant switch, old scope closed, logout.');
 }finally{await browser.close();await new Promise(resolve=>server.close(resolve));await deleteApp(admin);}
})().catch(e=>{console.error(e);process.exitCode=1;});

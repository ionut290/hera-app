// Real Firebase Auth + Firestore emulators; no live accounts or data.
const assert=require('node:assert/strict'),fs=require('node:fs'),http=require('node:http'),path=require('node:path');
const {chromium}=require('playwright');
const root=path.join(__dirname,'full-runtime-dist'),projectId='varga-cantieri-v2-staging';
(async()=>{
 assert.match(process.env.FIRESTORE_EMULATOR_HOST||'',/^127\.0\.0\.1:8080$/);
 assert.match(process.env.FIREBASE_AUTH_EMULATOR_HOST||'',/^127\.0\.0\.1:9099$/);
 const {initializeApp,deleteApp}=require('../node_modules/firebase-admin/app');
 const {getAuth}=require('../node_modules/firebase-admin/auth'),{getFirestore}=require('../node_modules/firebase-admin/firestore');
 const admin=initializeApp({projectId}),db=getFirestore(admin),auth=getAuth(admin);
 const uid='local-admin',email='local-admin@example.test',password='Local-Test-Only-42';
 await auth.createUser({uid,email,password,emailVerified:true,displayName:'Local test admin'});
 for(const id of ['demo-a','demo-b']){
  await db.doc('tenants/'+id).set({demo:true,status:'active',name:id,originalCreateEnabled:true,originalManualEnabled:true,originalArchivesEnabled:true,priceListEnabled:true,priceListWriteEnabled:true,snowCreateEnabled:true});
  await db.doc('tenants/'+id+'/members/'+uid).set({role:'admin',status:'active'});
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
  await page.goto('http://127.0.0.1:4173/');
  await page.locator('#email').fill(email);await page.locator('#password').fill(password);await page.locator('#login button').click();
  await page.locator('[data-tenant="demo-a"]').click();
  const frame=page.frameLocator('#runtime');
  await frame.locator('#side-menu').waitFor({state:'attached'});
  await page.waitForFunction(()=>window.__vargaLocalRuntime().session.session.context.tenantId==='demo-a');
  const a=await page.evaluate(async()=>{const s=window.__vargaLocalRuntime().session;window.__oldLocalScope=s;return (await s.session.firestore.collection('commesse').doc('same-id').get({source:'server'})).data().nome;});
  assert.equal(a,'demo-a original');
  // Writes from the actual iframe realm, using the selected company Firestore facade.
  const child=page.frames().find(f=>f.url().endsWith('/original.html'));assert(child);
  await child.evaluate(async()=>{const s=window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__;const stamp=window.firebase.firestore.FieldValue.serverTimestamp();await s.session.firestore.collection('commesse').add({nome:'Created from actual frame',codice:'',excelModelVersion:2,priceListVersion:2,workItemsModelVersion:2,percentualeRibassoGenerale:0.01,excelModelActivatedAt:stamp,excelModelActivatedBy:s.auth.currentUser.email,nextImpiantoNumber:1,parentCommessaId:null,creatoDa:s.auth.currentUser.email,createdAt:stamp});});
  assert.equal((await db.collection('tenants/demo-a/commesse').get()).size,2);
  await page.locator('[data-tenant="demo-b"]').click();
  await page.waitForFunction(()=>window.__vargaLocalRuntime().session.session.context.tenantId==='demo-b');
  assert.equal(await page.evaluate(async()=>{const s=window.__vargaLocalRuntime().session;return (await s.session.firestore.collection('commesse').doc('same-id').get({source:'server'})).data().nome;}),'demo-b original');
  assert.equal(await page.evaluate(()=>{try{window.__oldLocalScope.session.firestore.collection('commesse');return false;}catch{return true;}}),true);
  assert.equal((await db.collection('tenants/demo-b/commesse').get()).size,1);
  await page.locator('#logout').click();await page.locator('#login').waitFor({state:'visible'});
  assert.equal(await page.locator('#runtime').isHidden(),true);
  assert.equal(external.length,0,'Unexpected external requests: '+external.join(','));
  console.log('FULL PACKAGE REAL AUTH E2E: login, verified membership, original UI, frame write, tenant switch, old scope closed, logout.');
 }finally{await browser.close();await new Promise(resolve=>server.close(resolve));await deleteApp(admin);}
})().catch(e=>{console.error(e);process.exitCode=1;});

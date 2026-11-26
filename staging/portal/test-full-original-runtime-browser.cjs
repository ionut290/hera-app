// Full original HTML + candidate source startup smoke. Local-only fixture, no production SDK,
// no deployment, no auth tokens. Reports runtime errors as blockers rather than hiding them.
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');
const {chromium}=require('playwright');
(async()=>{
 const root=path.resolve(__dirname,'../..'),original=fs.readFileSync(path.join(root,'index.html'),'utf8');
 const {buildIntegratedRuntimeCandidate}=await import('../original-runtime-source.mjs');
 const candidate=buildIntegratedRuntimeCandidate(fs.readFileSync(path.join(root,'app.js'),'utf8'));
 const browser=await chromium.launch({headless:true});
 const origin='http://varga-runtime.test';
 try{
  const context=await browser.newContext({viewport:{width:390,height:844},serviceWorkers:'block'});
  const external=[],errors=[],served=new Set();
  await context.addInitScript(()=>{
   const storage=()=>{const m=new Map();return {getItem:k=>m.get(k)??null,setItem:(k,v)=>m.set(k,String(v)),removeItem:k=>m.delete(k),clear:()=>m.clear()};};
   const uid='smoke-operator',projectId='varga-cantieri-v2-staging',app={name:'fixture',options:{projectId}};
   const user={uid,email:'smoke@example.test',displayName:'Operatore demo'};
   const auth={app,currentUser:user,onAuthStateChanged(){return()=>{};}};
   const firestore={collection(){throw Error('Runtime smoke: database operations blocked');}};
   window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__={projectId,auth,session:{context:{uid,tenantId:'demo-smoke',role:'employee'},firestore},isActive:()=>true,localStorage:storage(),sessionStorage:storage(),approval:{verify:async()=>({allowed:true,profile:{uid,role:'user'}})}};
   window.firebase={initializeApp(){throw Error('Unscoped Firebase blocked');},auth(){throw Error('Unscoped auth blocked');},firestore(){throw Error('Unscoped Firestore blocked');}};
   window.__VARGA_STAGING_FIREBASE_CONFIG__={apiKey:'fixture-only'};
  });
  await context.route('**/*',async route=>{
   const url=new URL(route.request().url());
   if(url.origin!==origin){external.push(url.hostname);return route.abort();}
   if(route.request().method()!=='GET')return route.abort();
   const rel=decodeURIComponent(url.pathname).replace(/^\/+/,'')||'index.html',file=path.resolve(root,rel);
   if(!file.startsWith(root+path.sep)||!fs.existsSync(file)||!fs.statSync(file).isFile())return route.fulfill({status:404,body:''});
   if(rel==='index.html'){
    // Only inside this test, enable local scripts/styles. Network remains intercepted.
    const html=original.replace("default-src 'none'; base-uri 'none'; form-action 'none'","default-src 'none'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self'; connect-src 'none'; base-uri 'none'; form-action 'none'");
    assert.notEqual(html,original,'Expected staging quarantine marker');
    return route.fulfill({contentType:'text/html',body:html});
   }
   if(rel==='app.js'){
    served.add(rel);
    // Prove actual source reaches its original SDK bootstrap; do not claim full startup
    // until all external library/legacy module dependencies have been adapted.
    const marker='const PERSISTED_SESSION_KEY';
    const i=candidate.indexOf(marker);assert(i>0);
    const instrumented=candidate.slice(0,i)+'window.__vargaRuntimeBootstrap = {bound: db === window.__VARGA_ISOLATED_RUNTIME_CANDIDATE__.session.firestore, error: firebaseInitError ? String(firebaseInitError.message) : null};\n'+candidate.slice(i);
    return route.fulfill({contentType:'application/javascript',body:instrumented});
   }
   const ext=path.extname(file),types={'.js':'application/javascript','.css':'text/css','.html':'text/html','.json':'application/json','.webmanifest':'application/manifest+json','.png':'image/png','.svg':'image/svg+xml'};
   served.add(rel);return route.fulfill({contentType:types[ext]||'application/octet-stream',body:fs.readFileSync(file)});
  });
  const page=await context.newPage();page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(10000);
  await page.goto(origin+'/',{waitUntil:'domcontentloaded'});
  for(const id of ['app-startup-loading','auth-gate','side-menu','open-panel-commesse','open-panel-squadre','open-hours-btn','open-urban-furniture-btn','green-assistant-overlay']){
   assert.equal(await page.locator('#'+id).count(),1,'Original shell missing '+id);
  }
  assert(served.has('app.js'),'Complete candidate not loaded');
  const bootstrap=await page.evaluate(()=>window.__vargaRuntimeBootstrap);
  assert.deepEqual(bootstrap,{bound:true,error:null});
  console.log('FULL ORIGINAL SHELL: original HTML, original asset scripts, candidate SDK bootstrap verified.');
  console.log('Remaining browser runtime errors (diagnostic): '+JSON.stringify([...new Set(errors)].slice(0,30)));
  console.log('Blocked external hosts: '+JSON.stringify([...new Set(external)].sort()));
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});

'use strict';
// Browser smoke of original index.html + all local scripts. No live credentials or external network.
// A dedicated ephemeral browser profile, no SW, emulator-only Firebase config.
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');
const {chromium}=require('playwright');
const {createLocalAdapter}=require('./auth-adapter.emulator.cjs');
const {startLocalCompanyHttp}=require('./company-http.emulator.cjs');
const {randomUUID}=require('node:crypto');
const root=path.resolve(__dirname,'../..'),deps=process.env.VARGA_BROWSER_DEPS;
assert.ok(deps&&path.isAbsolute(deps)&&fs.existsSync(path.join(deps,'firebase','firebase-app-compat.js')),'Isolated browser SDK unavailable');
const origin='http://127.0.0.1:4173',project='demo-varga-multi-tenant';
const types={'.html':'text/html','.js':'application/javascript','.css':'text/css','.json':'application/json','.webmanifest':'application/manifest+json','.png':'image/png','.jpg':'image/jpeg','.svg':'image/svg+xml','.webp':'image/webp','.ico':'image/x-icon','.woff2':'font/woff2'};
const sdkNames=new Set(['firebase-app-compat.js','firebase-auth-compat.js','firebase-firestore-compat.js','firebase-storage-compat.js','firebase-functions-compat.js','firebase-messaging-compat.js']);
function localFile(p){
  const decoded=decodeURIComponent(p);if(decoded.includes('\\')||decoded.split('/').some(s=>s==='..'||s.startsWith('.')))return null;
  let rel=decoded.replace(/^\//,'')||'index.html';
  if(rel.includes('/')&&!/^(icons|piano-neve)\//.test(rel))return null;
  const full=path.resolve(root,rel);if(!full.startsWith(root+path.sep)||!types[path.extname(full)]||!fs.existsSync(full)||!fs.statSync(full).isFile())return null;
  if(!fs.realpathSync(full).startsWith(root+path.sep))return null;
  return full;
}
async function main(){
  assert.equal(process.env.FIRESTORE_EMULATOR_HOST,'127.0.0.1:8180');
  assert.equal(process.env.FIREBASE_AUTH_EMULATOR_HOST,'127.0.0.1:9199');
  const adapter=createLocalAdapter();
  const http=await startLocalCompanyHttp(adapter.dispatch);
  const password='Demo-'+randomUUID(),email='browser-'+randomUUID()+'@example.invalid';
  const user=await adapter.auth.createUser({email,password,emailVerified:true,displayName:'Synthetic browser user'});
  const db=adapter.db,seed=(p,d)=>db.doc(p).set(d);
  await Promise.all([
    seed('platformUsers/'+user.uid,{uid:user.uid,email,displayName:'Synthetic browser user',statoAccount:'attivo',accountStatus:'attivo',role:'user',ruolo:'user',banned:false,tenantLabSynthetic:true}),
    seed('tenants/browser-a',{name:'Browser company A',status:'active'}),
    seed('tenants/browser-b',{name:'Browser company B',status:'active'}),
    seed('tenants/browser-c',{name:'Unauthorized C',status:'active'}),
    seed('tenants/browser-a/members/'+user.uid,{role:'owner',status:'active'}),
    seed('tenants/browser-b/members/'+user.uid,{role:'owner',status:'active'}),
    seed('users/'+user.uid+'/memberships/browser-a',{status:'active'}),
    seed('users/'+user.uid+'/memberships/browser-b',{status:'active'}),
    seed('tenants/browser-a/commesse/same-id',{nome:'Browser job A',createdAt:new Date()}),
    seed('tenants/browser-b/commesse/same-id',{nome:'Browser job B',createdAt:new Date()}),
    seed('tenants/browser-c/commesse/same-id',{nome:'Private C',createdAt:new Date()})
  ]);
  let browser,context;
  try{
  browser=await chromium.launch({headless:true,args:['--disable-background-networking']});
  context=await browser.newContext({serviceWorkers:'block',viewport:{width:390,height:844},isMobile:true,hasTouch:true});
  const blocked=new Set(),errors=[],loaded=new Set();
    await context.routeWebSocket('**/*',ws=>{ws.close();});
    await context.route('**/*',async route=>{
      const u=new URL(route.request().url());
      if(u.origin==='http://127.0.0.1:8180'||u.origin==='http://127.0.0.1:9199')return route.continue();
      if(u.origin===origin&&u.pathname==='/api/v2/companies'){
        const request=route.request();
        if(request.method()!=='POST'||!request.headers().authorization)return route.fulfill({status:403,body:'{}'});
        const response=await fetch(http.origin+'/api/v2/companies',{method:'POST',headers:{origin:http.origin,'content-type':'application/json',authorization:request.headers().authorization},body:request.postData(),redirect:'error'});
        return route.fulfill({status:response.status,contentType:'application/json',body:await response.text()});
      }
      let f=null;
      if(u.origin===origin)f=localFile(u.pathname);
      else if(u.hostname==='www.gstatic.com'&&sdkNames.has(path.basename(u.pathname)))f=path.join(deps,'firebase',path.basename(u.pathname));
      else if(u.hostname==='unpkg.com'&&/^\/leaflet@1\.9\.4\/dist\/(leaflet\.(js|css)|images\/[a-zA-Z0-9_.-]+)$/.test(u.pathname))f=path.join(deps,'leaflet','dist',u.pathname.split('/dist/')[1]);
      if(!f||!fs.existsSync(f)){blocked.add(u.origin===origin?u.pathname:u.origin);return route.abort('blockedbyclient');}
      let body=fs.readFileSync(f);
      if(f===path.join(root,'firebase-config.js')){
        let s=body.toString('utf8');
        assert.match(s,/window\.firebaseConfig\s*=\s*\{[^\n]*\};/);
        s=s.replace(/window\.firebaseConfig\s*=\s*\{[^\n]*\};/,`window.firebaseConfig={apiKey:"demo-key",authDomain:"demo-varga-multi-tenant.firebaseapp.com",projectId:"demo-varga-multi-tenant",appId:"demo-app"};firebase.initializeApp(window.firebaseConfig);firebase.auth().useEmulator("http://127.0.0.1:9199",{disableWarnings:true});firebase.firestore().useEmulator("127.0.0.1",8180);window.VargaTenantLabEnabled=true;`);
        body=Buffer.from(s);
      }
      loaded.add(path.basename(f));
      return route.fulfill({status:200,contentType:types[path.extname(f)]||'application/octet-stream',body,headers:{'cache-control':'no-store'}});
    });
    const page=await context.newPage();page.on('pageerror',e=>errors.push(String(e.message).slice(0,200)));
    await page.goto(origin+'/index.html',{waitUntil:'domcontentloaded',timeout:45000});
    await page.waitForFunction(()=>window.firebase?.apps?.length>0&&firebase.app().options.projectId==='demo-varga-multi-tenant',null,{timeout:30000});
    await page.waitForFunction(()=>{const e=document.getElementById('auth-gate');return e&&!e.hidden&&!e.classList.contains('hidden')&&getComputedStyle(e).display!=='none'&&e.getBoundingClientRect().height>0;},null,{timeout:35000});
    assert.ok(loaded.has('app.js'),'Original app.js was not loaded');
    assert.ok(loaded.has('auth-login-fix.js'),'Original login runtime was not loaded');
    assert.ok(await page.locator('#commesse-lista').count(),'Original home markup absent');
    assert.equal(await page.evaluate(()=>firebase.auth().currentUser===null),true);
    assert.equal(await page.locator('#varga-company-open').isVisible(),false);
    assert.equal(await page.locator('#varga-tenant-admin-open').isVisible(),false);
    assert.ok(await page.evaluate(()=>typeof window.VargaTenantSession?.create==='function'));
    console.log(JSON.stringify({originalBrowserBoot:'passed',blockedDestinations:[...blocked].sort(),pageErrors:errors}));
    // Real synthetic email/password login through the existing app form.
    await page.locator('#auth-email-input').fill(email);
    await page.locator('#auth-password-input').fill(password);
    await page.locator('#auth-email-login-btn').click();
    await page.waitForFunction(uid=>firebase.auth().currentUser?.uid===uid,user.uid,{timeout:35000});
    await page.waitForFunction(()=>{const e=document.getElementById('access-approval-screen');return document.body.classList.contains('auth-authenticated')|| (e&&e.classList.contains('hidden')&&document.getElementById('auth-gate')?.classList.contains('hidden'));},null,{timeout:35000});
    await page.waitForFunction(()=>{const b=document.getElementById('varga-company-open');return b&&!b.hidden&&!b.classList.contains('hidden');},null,{timeout:35000});
    await page.locator('#menu-toggle-btn').click();
    await page.locator('#varga-company-open').click();
    const list=page.locator('#varga-company-list');
    await list.getByRole('button',{name:'Browser company A · owner',exact:true}).waitFor();
    assert.equal(await list.getByRole('button',{name:/Unauthorized C/}).count(),0);
    await list.getByRole('button',{name:'Browser company A · owner',exact:true}).click();
    await page.locator('#varga-company-ordinary').click();
    const jobs=page.locator('#varga-company-jobs');
    await jobs.getByText('Browser job A',{exact:true}).waitFor();
    assert.ok(!(await jobs.innerText()).includes('Browser job B'));
    await list.getByRole('button',{name:'Browser company B · owner',exact:true}).click();
    await page.locator('#varga-company-ordinary').click();
    await jobs.getByText('Browser job B',{exact:true}).waitFor();
    assert.ok(!(await jobs.innerText()).includes('Browser job A'));
    await page.evaluate(()=>firebase.auth().signOut());
    await page.waitForFunction(()=>document.getElementById('varga-company-open').hidden&&document.getElementById('varga-company-jobs').children.length===0);
    console.log(JSON.stringify({syntheticOriginalLoginAndSelector:'passed',pageErrors:errors}));
    // Errors remain diagnostic; full operational runtime and native login are not covered.
  }finally{if(context)await context.close();if(browser)await browser.close();await http.close();await adapter.close();}
}
main().catch(e=>{console.error(e);process.exitCode=1;});

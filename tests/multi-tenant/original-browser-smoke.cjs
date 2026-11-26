'use strict';
// Browser smoke of original index.html + all local scripts. No live credentials or external network.
// A dedicated ephemeral browser profile, no SW, emulator-only Firebase config.
const assert=require('node:assert/strict'),fs=require('node:fs'),path=require('node:path');
const {chromium}=require('playwright');
const {createLocalAdapter}=require('./auth-adapter.emulator.cjs');
const {startLocalCompanyHttp}=require('./company-http.emulator.cjs');
const {randomUUID}=require('node:crypto');
const httpNode=require('node:http');
const root=path.resolve(__dirname,'../..'),deps=process.env.VARGA_BROWSER_DEPS;
assert.ok(deps&&path.isAbsolute(deps)&&fs.existsSync(path.join(deps,'firebase','firebase-app-compat.js')),'Isolated browser SDK unavailable');
const origin='http://127.0.0.1:4173',project='demo-varga-multi-tenant';
const FS_RPC=/^\/(?:google\.firestore\.v1\.Firestore\/(?:Listen|Write)\/channel|v1\/projects\/demo-varga-multi-tenant\/databases\/\(default\)\/documents(?::[A-Za-z]+|\/.*)?|google\.firestore\.v1\.Firestore\/[A-Za-z]+)$/;
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
async function startFirestoreProxy(metrics){
  const server=httpNode.createServer((req,res)=>{
    const u=new URL(req.url,origin);
    if(req.headers.host!==new URL(origin).host||!FS_RPC.test(u.pathname)){
      res.writeHead(404);res.end();return;
    }
    const headers={...req.headers,host:'127.0.0.1:8180'};
    delete headers.connection;
    delete headers['proxy-connection'];
    // Keep real authorization and protocol parameters; do not grant any access.
    const upstream=httpNode.request({host:'127.0.0.1',port:8180,method:req.method,path:u.pathname+u.search,headers},r=>{
      metrics.push({path:u.pathname.split('/documents')[0],status:r.statusCode});
      const responseHeaders={...r.headers};
      delete responseHeaders.connection;
      res.writeHead(r.statusCode||502,responseHeaders);r.pipe(res);
    });
    upstream.on('error',()=>{if(!res.headersSent)res.writeHead(502);res.end();});
    res.on('close',()=>upstream.destroy());
    req.pipe(upstream);
  });
  await new Promise((resolve,reject)=>{server.once('error',reject);server.listen(4173,'127.0.0.1',resolve);});
  return {close:()=>new Promise(resolve=>{server.closeAllConnections();server.close(()=>resolve());})};
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
  // Independent check of the actual emulator account. Do not copy the resulting
  // token into the browser: the existing form must authenticate for itself.
  const credentialCheck=await fetch('http://127.0.0.1:9199/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=demo-key',{method:'POST',headers:{'content-type':'application/json'},body:JSON.stringify({email,password,returnSecureToken:true}),redirect:'error'});
  const credentialResult=await credentialCheck.json();
  assert.ok(credentialCheck.ok&&credentialResult.localId===user.uid&&credentialResult.idToken,'Synthetic Auth emulator account is not usable');
  let browser,context,page,proxy,phase='setup';
  const firestoreResponses=[];
  const blocked=new Set(),errors=[],loaded=new Set(),authRequests=[],authResponses=[],authConsole=[];
  const mark=p=>{phase=p;console.log('Browser phase: '+p);};
  try{
  proxy=await startFirestoreProxy(firestoreResponses);
  browser=await chromium.launch({headless:true,args:['--disable-background-networking']});
  context=await browser.newContext({serviceWorkers:'block',viewport:{width:390,height:844},isMobile:true,hasTouch:true});

    await context.routeWebSocket('**/*',ws=>{ws.close();});
    await context.route('**/*',async route=>{
      const u=new URL(route.request().url());
      if(u.origin===origin&&FS_RPC.test(u.pathname))return route.continue();
      if(u.origin==='http://127.0.0.1:8180')return route.abort('blockedbyclient');
      // Same-origin transport to the actual Auth emulator avoids cross-origin
      // loopback restrictions. No fake responses, tokens or authentication.
      if(u.origin===origin&&/^\/(identitytoolkit|securetoken)\.googleapis\.com\//.test(u.pathname)){
        const request=route.request(),method=request.method();
        if(!['GET','POST'].includes(method))return route.abort('blockedbyclient');
        const target=new URL(u.pathname+u.search,'http://127.0.0.1:9199');
        authRequests.push({path:u.pathname,method});
        const headers={};
        for(const name of ['content-type','authorization','x-client-version','x-firebase-gmpid'])if(request.headers()[name])headers[name]=request.headers()[name];
        const response=await fetch(target,{method,headers,body:method==='POST'?request.postDataBuffer():undefined,redirect:'error'});
        const raw=await response.text();
        let code=null;
        if(!response.ok){try{code=JSON.parse(raw)?.error?.message||null;}catch{}}
        authResponses.push({path:u.pathname,status:response.status,code:typeof code==='string'?code.slice(0,100):null});
        return route.fulfill({status:response.status,contentType:response.headers.get('content-type')||'application/json',body:raw,headers:{'cache-control':'no-store'}});
      }
      if(u.origin==='http://127.0.0.1:9199')return route.abort('blockedbyclient');
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
        s=s.replace(/window\.firebaseConfig\s*=\s*\{[^\n]*\};/,`window.firebaseConfig={apiKey:"demo-key",authDomain:"demo-varga-multi-tenant.firebaseapp.com",projectId:"demo-varga-multi-tenant",appId:"demo-app"};firebase.initializeApp(window.firebaseConfig);firebase.auth().useEmulator("http://127.0.0.1:4173",{disableWarnings:true});firebase.firestore().useEmulator("127.0.0.1",4173);window.VargaTenantLabEnabled=true;`);
        body=Buffer.from(s);
      }
      loaded.add(path.basename(f));
      return route.fulfill({status:200,contentType:types[path.extname(f)]||'application/octet-stream',body,headers:{'cache-control':'no-store'}});
    });
    page=await context.newPage();page.on('pageerror',e=>errors.push(String(e.message).slice(0,200)));
    page.on('console',message=>{
      const value=message.text();
      // Keep diagnostic categories only. No account data, console objects or tokens.
      if(/Errore login email\/password|AUTH READY|USER LOGGED|USER NOT LOGGED|Access verification failed|Sessione salvata non valida|Persistenza Firebase LOCAL/.test(value)){
        const category=value.match(/Errore login email\/password|AUTH READY|USER LOGGED|USER NOT LOGGED|Access verification failed|Sessione salvata non valida|Persistenza Firebase LOCAL/)?.[0];
        const code=value.match(/\b(?:auth|firestore)\/[a-z-]+\b/)?.[0]||null;
        authConsole.push({category,code});
      }
    });
    await page.addInitScript(()=>{
      window.__vargaBrowserLabFormSubmits=0;
      document.addEventListener('submit',e=>{if(e.target?.id==='auth-email-form')window.__vargaBrowserLabFormSubmits++;},true);
    });
    mark('original boot');
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
    mark('original email login');
    await page.locator('#auth-email-input').fill(email);
    await page.locator('#auth-password-input').fill(password);
    assert.equal(await page.locator('#auth-email-form').evaluate(f=>f.checkValidity()),true,'Original form did not accept synthetic credentials');
    await page.locator('#auth-email-login-btn').click();
    await page.waitForFunction(()=>window.__vargaBrowserLabFormSubmits>0,null,{timeout:10000});
    mark('waiting for authenticated Firebase identity');
    await page.waitForFunction(uid=>firebase.auth().currentUser?.uid===uid,user.uid,{timeout:35000});
    assert.ok(authResponses.some(x=>x.path.includes('signInWithPassword')&&x.status===200),'No real emulator sign-in response observed');
    mark('waiting for original account approval');
    // First prove real server-side profile read using the signed-in SDK; no mock or
    // profile bootstrap injection. The original approval gate must also finish.
    const profileCheck=await page.evaluate(async uid=>{
      try{
        const d=await firebase.firestore().collection('platformUsers').doc(uid).get({source:'server'});
        return {exists:d.exists,approved:d.data()?.accountStatus==='attivo',synthetic:d.data()?.tenantLabSynthetic===true};
      }catch(e){return {errorCode:String(e?.code||'unknown')};}
    },user.uid);
    assert.deepEqual(profileCheck,{exists:true,approved:true,synthetic:true},'Original synthetic profile is not readable from the Firestore server');
    await page.waitForFunction(()=>{
      const e=document.getElementById('access-approval-screen'),gate=document.getElementById('auth-gate'),home=document.getElementById('home-page');
      return e&&e.classList.contains('hidden')&&gate&&gate.classList.contains('hidden')&&home&&!home.hidden&&!home.classList.contains('hidden')&&!['auth-pending','auth-required','auth-banned'].some(c=>document.body.classList.contains(c));
    },null,{timeout:35000});
    await page.waitForFunction(()=>{const b=document.getElementById('varga-company-open');return b&&!b.hidden&&!b.classList.contains('hidden');},null,{timeout:35000});
    mark('authorized company selector');
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
    mark('switch company');
    await list.getByRole('button',{name:'Browser company B · owner',exact:true}).click();
    await page.locator('#varga-company-ordinary').click();
    await jobs.getByText('Browser job B',{exact:true}).waitFor();
    assert.ok(!(await jobs.innerText()).includes('Browser job A'));
    mark('self-service registration and member invitation');
    await page.locator('#varga-company-register-name').fill('Browser starter company');
    await page.locator('#varga-company-register button[type=submit]').click();
    await page.locator('#varga-company-name').getByText('Browser starter company',{exact:true}).waitFor();
    await page.locator('#varga-company-plan-summary').getByText(/Commesse 0\/1.*Impianti 0\/10.*Admin 1\/1.*Operatori 0\/3/).waitFor();
    await page.locator('#varga-company-members-refresh').click();
    await page.locator('#varga-company-members-list').getByText(/Titolare/).waitFor();
    await page.locator('#varga-company-invite-email').fill('invited-browser@example.invalid');
    await page.locator('#varga-company-invite-form button[type=submit]').click();
    await page.locator('#varga-company-invite-result').waitFor();
    const invitation=await page.locator('#varga-company-invite-code').inputValue();
    assert.ok(/^[A-Za-z0-9_-]{1,128}:[a-f0-9]{64}$/.test(invitation),'Invitation code format invalid');
    await page.locator('#varga-company-invite-hide').click();
    assert.equal(await page.locator('#varga-company-invite-code').inputValue(),'');
    console.log('Browser starter registration and invitation passed');
    mark('logout');
    await page.evaluate(()=>firebase.auth().signOut());
    await page.waitForFunction(()=>document.getElementById('varga-company-open').hidden&&document.getElementById('varga-company-jobs').children.length===0);
    mark('first tenant access without legacy profile');
    const freshEmail='tenant-only-'+randomUUID()+'@example.invalid',freshPassword='Demo-'+randomUUID();
    const fresh=await adapter.auth.createUser({email:freshEmail,password:freshPassword,emailVerified:true});
    assert.equal((await db.doc('platformUsers/'+fresh.uid).get()).exists,false);
    await page.reload({waitUntil:'domcontentloaded'});
    await page.locator('#auth-email-input').fill(freshEmail);
    await page.locator('#auth-password-input').fill(freshPassword);
    await page.locator('#auth-email-login-btn').click();
    await page.waitForFunction(uid=>firebase.auth().currentUser?.uid===uid,fresh.uid,{timeout:35000});
    await page.locator('#access-approval-screen').waitFor({state:'visible',timeout:35000});
    await page.locator('#varga-company-entry-open').click();
    await page.locator('#varga-company-dialog').waitFor({state:'visible'});
    await page.locator('#varga-company-register-name').fill('First autonomous company');
    await page.locator('#varga-company-register button[type=submit]').click();
    await page.locator('#varga-company-name').getByText('First autonomous company',{exact:true}).waitFor();
    await page.locator('#varga-company-plan-summary').getByText(/Commesse 0\/1.*Impianti 0\/10.*Admin 1\/1.*Operatori 0\/3/).waitFor();
    assert.equal((await db.doc('platformUsers/'+fresh.uid).get()).exists,false,'Legacy profile unexpectedly created');
    assert.equal(await page.evaluate(()=>document.body.classList.contains('access-approval-locked')),true,'Legacy gate must remain locked');
    await page.locator('#varga-company-close').click();
    await page.locator('#access-approval-screen').waitFor({state:'visible'});
    await page.evaluate(()=>firebase.auth().signOut());
    await page.waitForFunction(()=>document.getElementById('varga-company-entry').hidden);
    console.log('Tenant-only first access passed without granting legacy approval');
    console.log(JSON.stringify({syntheticOriginalLoginAndSelector:'passed',pageErrors:errors}));
    // Errors remain diagnostic; full operational runtime and native login are not covered.
  }catch(error){
    let loginState=null;
    if(page&&!page.isClosed())try{loginState=await page.evaluate(()=>{
      const feedback=document.getElementById('auth-email-feedback')?.textContent||'';
      const category=/Login completato/i.test(feedback)?'success':/Accesso in corso/i.test(feedback)?'pending':/Email o password non corretta/i.test(feedback)?'invalid-credential':/Connessione non disponibile/i.test(feedback)?'network':/Accesso non riuscito/i.test(feedback)?'failed':/Inserisci email/i.test(feedback)?'missing-input':feedback?'other':'empty';
      return {formSubmits:window.__vargaBrowserLabFormSubmits||0,feedbackCategory:category,buttonDisabled:document.getElementById('auth-email-login-btn')?.disabled,authPresent:!!window.firebase?.auth?.().currentUser,authVerified:window.firebase?.auth?.().currentUser?.emailVerified===true,gateHidden:document.getElementById('auth-gate')?.classList.contains('hidden'),approvalHidden:document.getElementById('access-approval-screen')?.classList.contains('hidden')};
    });}catch{}
    console.error(JSON.stringify({failedPhase:phase,loginState,authRequests,authResponses,authConsole,firestoreResponses,loadedAssets:[...loaded].sort(),blockedDestinations:[...blocked].sort(),pageErrors:errors}));
    throw error;
  }finally{if(context)await context.close();if(browser)await browser.close();if(proxy)await proxy.close();await http.close();await adapter.close();}
}
main().catch(e=>{console.error(e);process.exitCode=1;});

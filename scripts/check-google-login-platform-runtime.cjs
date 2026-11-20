"use strict";
// Executes the REAL login handler against synthetic Firebase/Capacitor adapters.
// Never signs in to a real account, opens an OAuth popup, or calls Firebase.
const assert=require("node:assert/strict");
const fs=require("node:fs");
const vm=require("node:vm");
const source=fs.readFileSync("auth-login-fix.js","utf8");
const start=source.indexOf("(function installGoogleLoginFix() {");
const end=source.indexOf("(function installLoginPasswordVisibility()",start);
assert.ok(start>=0&&end>start,"Google login module missing");
const moduleSource=source.slice(start,end);
const capacitor=JSON.parse(fs.readFileSync("capacitor.config.json","utf8"));
assert.ok(capacitor.plugins.FirebaseAuthentication.providers.includes("google.com"));
assert.equal(capacitor.plugins.FirebaseAuthentication.skipNativeAuth,true);
const iosSetup=fs.readFileSync("scripts/configure-ios-firebase.py","utf8");
assert.match(iosSetup,/REVERSED_CLIENT_ID/);
assert.match(iosSetup,/CFBundleURLSchemes/);
const html=fs.readFileSync("index.html","utf8");
const sw=fs.readFileSync("sw.js","utf8");
const version=/auth-login-fix\.js\?v=([0-9a-zA-Z._-]+)/.exec(html)?.[1];
assert.ok(version&&sw.includes("auth-login-fix.js?v="+version),"PWA cache must match current Google login fix");

function make({platform="ios",plugin=true,idToken="synthetic-id-token",nativeError=null,
             popupError=null,pendingNative=false}={}){
  const calls=[],alerts=[],errors=[],listeners={};
  let resolveNative;
  const items=new Map();
  function element(id){return {id,hidden:false,disabled:false,tabIndex:0,textContent:"",
    removeAttribute(){},setAttribute(){},focus(){calls.push("focus-email");}};}
  for(const id of ["auth-gate-login-btn","auth-email-input","auth-email-login-btn","auth-gate-message"])items.set(id,element(id));
  const divider=element("divider");
  const document={
    readyState:"complete",documentElement:{classList:{toggle(name,on){calls.push(["class",name,on]);}}},
    getElementById(id){return items.get(id)||null;},
    querySelector(value){assert.equal(value,".auth-gate-divider");return divider;},
    addEventListener(name,handler,capture){listeners[name]=handler;assert.equal(capture,true);}
  };
  const currentAuth={
    onAuthStateChanged(next){return ()=>{};},
    async signInWithCredential(credential){calls.push(["firebase-credential",credential]);return {user:{uid:"synthetic-user"}};},
    async signInWithPopup(provider){
      calls.push(["web-popup",provider.providerId]);
      if(popupError)throw popupError;
      return {user:{uid:"synthetic-web-user"}};
    }
  };
  const auth=()=>currentAuth;
  auth.GoogleAuthProvider=class GoogleAuthProvider{
    constructor(){this.providerId="google.com";}
    addScope(scope){assert.equal(scope,"https://www.googleapis.com/auth/userinfo.email");}
    setCustomParameters(obj){assert.equal(obj.prompt,"select_account");}
    static credential(token,accessToken){
      assert.equal(token,idToken);
      return {providerId:"google.com",idToken:token};
    }
  };
  const firebase={apps:[{}],auth};
  const nativePlugin={
    async signInWithGoogle(options){
      assert.deepEqual(JSON.parse(JSON.stringify(options)),{skipNativeAuth:true});
      calls.push("native-google");
      if(nativeError)throw nativeError;
      if(pendingNative)await new Promise(resolve=>{resolveNative=resolve;});
      return {credential:{idToken}};
    }
  };
  const window={
    firebase,
    Capacitor:{isNativePlatform:()=>platform!=="web",getPlatform:()=>platform,
      Plugins:plugin?{FirebaseAuthentication:nativePlugin}:{}
    },
    setTimeout(){throw Error("Unexpected wait for Firebase: test preloads Firebase");}
  };
  vm.runInNewContext(moduleSource,{
    window,document,firebase,alert:m=>alerts.push(String(m)),
    console:{error:e=>errors.push(String(e)),warn(){}},
    queueMicrotask:fn=>fn(),Promise,Map,Set,Date,Number,Object,Array,String,Math
  },{filename:"auth-login-fix.js"});
  function click(id="auth-gate-login-btn"){
    const button=items.get(id);
    const event={target:{closest:()=>button},
      preventDefault(){},stopPropagation(){},stopImmediatePropagation(){}};
    return listeners.click(event);
  }
  return {calls,alerts,errors,window,document,button:items.get("auth-gate-login-btn"),click,
    release(){assert.equal(typeof resolveNative,"function");resolveNative();}};
}

(async()=>{
  const ios=make();
  assert.equal(ios.button.hidden,false);
  await ios.click();
  assert.equal(ios.calls.filter(x=>x==="native-google").length,1);
  assert.equal(ios.calls.filter(x=>Array.isArray(x)&&x[0]==="firebase-credential").length,1);
  assert.equal(ios.calls.filter(x=>Array.isArray(x)&&x[0]==="web-popup").length,0);
  assert.equal(ios.button.disabled,false);
  await ios.window.loginWithGoogle();
  assert.equal(ios.calls.filter(x=>x==="native-google").length,2);
  assert.equal(ios.calls.filter(x=>Array.isArray(x)&&x[0]==="web-popup").length,0);

  const web=make({platform:"web"});
  await web.click();
  assert.equal(web.calls.filter(x=>Array.isArray(x)&&x[0]==="web-popup").length,1);
  assert.equal(web.calls.filter(x=>x==="native-google").length,0);
  await web.window.loginWithGoogle();
  assert.equal(web.calls.filter(x=>Array.isArray(x)&&x[0]==="web-popup").length,2);

  const android=make({platform:"android"});
  assert.equal(android.button.hidden,true);
  await android.click();
  assert.ok(android.calls.includes("focus-email"));
  assert.equal(android.calls.filter(x=>Array.isArray(x)&&x[0]==="web-popup").length,0);
  await assert.rejects(android.window.loginWithGoogle(),/solo l'accesso con email e password/);

  const missing=make({plugin:false});
  await missing.click();
  assert.match(missing.alerts.at(-1),/Plugin Google non disponibile nell'app iPhone/);
  assert.ok(!missing.calls.includes("native-google"));
  assert.equal(missing.calls.filter(x=>Array.isArray(x)&&x[0]==="web-popup").length,0);

  const missingToken=make({idToken:null});
  await missingToken.click();
  assert.match(missingToken.alerts.at(-1),/token di accesso valido/);
  assert.equal(missingToken.calls.filter(x=>Array.isArray(x)&&x[0]==="firebase-credential").length,0);

  const cancelled=make({nativeError:{code:"12501",message:"Cancelled"}});
  await cancelled.click();
  assert.match(cancelled.alerts.at(-1),/annullato/);
  assert.equal(cancelled.calls.filter(x=>Array.isArray(x)&&x[0]==="web-popup").length,0);

  const racing=make({pendingNative:true});
  const first=racing.click();
  const second=racing.click();
  assert.equal(racing.calls.filter(x=>x==="native-google").length,1);
  racing.release();
  await Promise.all([first,second]);
  assert.equal(racing.calls.filter(x=>Array.isArray(x)&&x[0]==="firebase-credential").length,1);

  const failedDomain=make({platform:"web",popupError:{code:"auth/unauthorized-domain"}});
  await failedDomain.click();
  assert.match(failedDomain.alerts.at(-1),/Domini|Dominio.*Firebase/);

  console.log("PASS: iPhone uses Google native credential without WKWebView popup, web retains popup, Android email/password only, no silent fallback, errors and double tap safe.");
})().catch(error=>{console.error(error);process.exitCode=1;});

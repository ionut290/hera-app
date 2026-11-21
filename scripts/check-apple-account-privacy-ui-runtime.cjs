"use strict";
/*
 * Executes the REAL account-privacy-controls.js in a VM with synthetic DOM,
 * Firebase and Capacitor adapters. No real identity, server, network, IAP,
 * account deletion, enterprise data or native device is accessed.
 */
const assert=require("node:assert/strict");
const fs=require("node:fs");
const vm=require("node:vm");
const source=fs.readFileSync("account-privacy-controls.js","utf8");
const staged=fs.readFileSync("functions/account-deletion-request.js","utf8");
const entry=fs.readFileSync("functions/main.js","utf8");
assert.doesNotMatch(entry,/require\(["']\.\/account-deletion-request["']\)/);
assert.match(staged,/reviewBeforeErasureRequired:\s*true/);
assert.match(source,/window\.__VARGA_ACCOUNT_DELETION_REQUEST_ENABLED === true/);

function harness({platform="ios",request=false,sla=undefined,immediate=false,
                  user=true,plugin=true,response="pending",reasonCode=undefined,defer=false}={}){
  const elements=[],alerts=[],calls=[];
  let release;
  class Element {
    constructor(id=""){this.id=id;this.dataset={};this.children=[];this.parentElement=null;this.events={};this.disabled=false;elements.push(this);}
    appendChild(child){child.parentElement=this;this.children.push(child);return child;}
    insertAdjacentElement(position,child){assert.equal(position,"afterend");return this.parentElement.appendChild(child);}
    addEventListener(name,cb){this.events[name]=cb;}
    setAttribute(name,val){this[name]=val;}
    click(){assert.equal(typeof this.events.click,"function");return this.events.click();}
  }
  const root=new Element("root");
  const login=root.appendChild(new Element("login-parent"));
  login.appendChild(new Element("auth-gate-login-btn"));
  root.appendChild(new Element("user-details-panel"));
  const menu=root.appendChild(new Element("menu"));
  const pending=root.appendChild(new Element("pending"));
  pending.appendChild(new Element("access-approval-logout"));
  const document={
    readyState:"complete",
    createElement(){return new Element();},
    getElementById(id){return elements.find(el=>el.id===id&&el.parentElement)||null;},
    querySelector(q){assert.equal(q,"#side-menu .side-menu-body .menu-section:nth-child(2)");return menu;},
    querySelectorAll(q){
      const k={
        "[data-request-varga-deletion]":"requestVargaDeletion",
        "[data-check-varga-deletion-status]":"checkVargaDeletionStatus",
        "[data-delete-varga-account]":"deleteVargaAccount"
      }[q];
      assert.ok(k,"unknown selector: "+q);
      return elements.filter(el=>el.parentElement&&Object.prototype.hasOwnProperty.call(el.dataset,k));
    },
    addEventListener(){throw Error("DOM must be ready");}
  };
  const auth={currentUser:user?{uid:"synthetic-user",email:"synthetic@example.test",providerData:[]}:null,
    async signInWithCredential(c){calls.push({name:"apple-signin",payload:c});},
    async signInWithPopup(provider){
      assert.ok(provider instanceof firebase.auth.OAuthProvider);
      calls.push({name:"apple-web-popup"});
    }};
  const firebase={
    apps:[{}],
    auth:()=>auth,
    app:()=>({functions(region){
      assert.equal(region,"europe-west1");
      return {httpsCallable(name){
        assert.ok(["requestOwnVargaAccountDeletion","getOwnVargaAccountDeletionStatus"].includes(name),
          "Destructive/unapproved callable must NEVER be invoked in UI test");
        return async payload=>{
          calls.push({name,payload});
          if(defer&&name==="requestOwnVargaAccountDeletion"){
            await new Promise(resolve=>{release=resolve;});
          }
          return {data:name==="getOwnVargaAccountDeletionStatus"
            ?{requested:true,status:response,...(reasonCode===undefined?{}:{reasonCode})}
            :{requested:true,status:response,alreadyExists:false}};
        };
      }};
    }})
  };
  firebase.auth.OAuthProvider=class {
    constructor(name){assert.equal(name,"apple.com");}
    credential({idToken,rawNonce}){
      assert.equal(idToken,"synthetic-id-token");assert.equal(rawNonce,"synthetic-nonce");
      return {provider:"apple.com",idToken,rawNonce};
    }
  };
  const window={
    __VARGA_ACCOUNT_DELETION_ENABLED:immediate,
    __VARGA_ACCOUNT_DELETION_REQUEST_ENABLED:request,
    __VARGA_ACCOUNT_DELETION_REQUEST_SLA_DAYS:sla,
    Capacitor:{
      isNativePlatform:()=>platform!=="web",
      getPlatform:()=>platform,
      Plugins:plugin?{FirebaseAuthentication:{
        async signInWithApple(options){
          assert.equal(options.skipNativeAuth,true);
          return {credential:{idToken:"synthetic-id-token",nonce:"synthetic-nonce"}};
        }
      }}:{}
    },
    confirm:()=>true,
    prompt:()=>"RICHIEDO",
    location:{reload(){throw Error("Unexpected destructive reload");}}
  };
  // The production guard explicitly checks window.firebase before using Firebase Auth.
  window.firebase=firebase;
  vm.runInNewContext(source,{window,document,firebase,alert:t=>alerts.push(String(t)),
    console:{warn(){}},Promise,Date,Math,Number,Object,Array,Set,String},
    {filename:"account-privacy-controls.js"});
  return {
    get:id=>document.getElementById(id),
    all:q=>document.querySelectorAll(q),
    alerts,calls,release:()=>{assert.equal(typeof release,"function");release();}
  };
}

(async()=>{
  const normal=harness();
  assert.ok(normal.get("auth-apple-login-btn"));
  assert.equal(normal.all("[data-request-varga-deletion]").length,0);
  assert.equal(normal.all("[data-delete-varga-account]").length,0);
  await normal.get("auth-apple-login-btn").click();
  assert.equal(normal.calls[0].name,"apple-signin");
  assert.equal(normal.calls[0].payload.provider,"apple.com");

  const unavailable=harness({plugin:false});
  await unavailable.get("auth-apple-login-btn").click();
  assert.equal(unavailable.calls.length,0);
  assert.match(unavailable.alerts[0],/Login Apple nativo non disponibile/);
  const web=harness({platform:"web"});
  assert.ok(web.get("auth-apple-login-btn"),"Apple login must be discoverable on web");
  await web.get("auth-apple-login-btn").click();
  assert.equal(web.calls.length,1);
  assert.equal(web.calls[0].name,"apple-web-popup");
  assert.equal(harness({platform:"android"}).get("auth-apple-login-btn"),null,
    "Android native remains email/password only");
  const info=harness({platform:"web"});
  assert.ok(info.get("varga-deletion-info-menu"),
    "Users must see a truthful deletion status even without a backend");
  assert.ok(info.get("varga-deletion-info-profile"));
  assert.ok(info.get("varga-deletion-info-pending"));
  assert.ok(info.get("varga-deletion-info-gate"));
  assert.equal(info.get("varga-deletion-info-menu").href, "./delete-account.html",
    "Show the published account-deletion information/help page");
  assert.equal(info.alerts.length,0,
    "A passive help link must not claim any account has been deleted");
  assert.equal(info.calls.length,0,
    "Information-only control must never call Firebase or delete accounts");

  for(const cfg of [
    {request:false,sla:30},{request:true},
    {request:true,sla:0},{request:true,sla:91},{request:true,sla:30.5}
  ]){
    const app=harness(cfg);
    assert.equal(app.all("[data-request-varga-deletion]").length,0);
    assert.equal(app.all("[data-check-varga-deletion-status]").length,0);
  }

  const ready=harness({request:true,sla:30});
  assert.equal(ready.get("varga-deletion-info-menu"),null,
    "Informational fallback must disappear for approved request backend");
  assert.equal(ready.all("[data-request-varga-deletion]").length,4);
  assert.equal(ready.all("[data-check-varga-deletion-status]").length,4);
  assert.equal(ready.all("[data-delete-varga-account]").length,0);
  await ready.get("varga-deletion-request-profile").click();
  assert.deepEqual(JSON.parse(JSON.stringify(ready.calls[0])),{
    name:"requestOwnVargaAccountDeletion",
    payload:{confirmation:"RICHIEDO ELIMINAZIONE"}
  });
  assert.match(ready.alerts.at(-1),/NON è ancora eliminato/);
  await ready.get("varga-deletion-request-profile-status").click();
  assert.equal(ready.calls[1].name,"getOwnVargaAccountDeletionStatus");
  assert.deepEqual(JSON.parse(JSON.stringify(ready.calls[1].payload)),{});

  const race=harness({request:true,sla:30,defer:true});
  const a=race.get("varga-deletion-request-profile").click();
  const b=race.get("varga-deletion-request-menu").click();
  assert.equal(race.calls.filter(c=>c.name==="requestOwnVargaAccountDeletion").length,1);
  race.release();await Promise.all([a,b]);
  assert.equal(race.calls.length,1);

  const signedOut=harness({request:true,sla:30,user:false});
  await signedOut.get("varga-deletion-request-profile").click();
  assert.equal(signedOut.calls.length,0);
  assert.match(signedOut.alerts.at(-1),/prima accedere/);

  const already=harness({request:true,sla:30,response:"completed"});
  await already.get("varga-deletion-request-profile").click();
  assert.match(already.alerts.at(-1),/nessuna nuova cancellazione/i);

  const reason=harness({
    request:true,sla:30,response:"rejected",reasonCode:"retention_review_required"
  });
  await reason.get("varga-deletion-request-profile-status").click();
  assert.match(reason.alerts.at(-1),/obblighi di conservazione/);
  assert.match(reason.alerts.at(-1),/ripresentarla/);
  assert.match(reason.alerts.at(-1),/senza cancellazione/);

  const opaque=harness({
    request:true,sla:30,response:"rejected",
    reasonCode:"Sensitive private reviewer note - NEVER show"
  });
  await opaque.get("varga-deletion-request-profile-status").click();
  assert.match(opaque.alerts.at(-1),/motivazione non è disponibile/);
  assert.equal(opaque.alerts.at(-1).includes("Sensitive private reviewer note"),false);

  const noReason=harness({request:true,sla:30,response:"rejected"});
  await noReason.get("varga-deletion-request-profile-status").click();
  assert.match(noReason.alerts.at(-1),/motivazione non è disponibile/);

  const demo=harness({request:true,sla:30,immediate:true});
  assert.equal(demo.get("varga-deletion-info-menu"),null);
  assert.equal(demo.all("[data-request-varga-deletion]").length,0);
  assert.equal(demo.all("[data-delete-varga-account]").length,4);
  assert.equal(demo.calls.length,0);

  console.log("PASS: Apple web popup and native iOS login; Android excluded; visible non-destructive privacy information; deletion flags/SLA/requests remain gated (synthetic only).");
})().catch(e=>{console.error(e);process.exitCode=1;});

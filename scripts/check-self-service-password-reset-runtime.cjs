"use strict";
// Run the ACTUAL password-reset UI code with synthetic Firebase Auth and DOM.
// No mail is sent; no live account is used; no password is ever stored.
const assert=require("node:assert/strict");
const fs=require("node:fs");
const vm=require("node:vm");
const source=fs.readFileSync("first-login-password.js","utf8");
const html=fs.readFileSync("index.html","utf8");
const sw=fs.readFileSync("sw.js","utf8");
const ver=html.match(/first-login-password\.js\?v=([A-Za-z0-9._-]+)/)?.[1];
assert.ok(ver&&sw.includes("first-login-password.js?v="+ver),"PWA cache must carry same reset UI build");
for(const required of [
  'id="auth-request-password-btn"',
  'id="auth-password-reset-dialog"',
  'id="auth-password-reset-form"',
  'id="auth-password-reset-email"',
  'id="auth-password-reset-feedback"',
  'id="auth-password-reset-submit"',
  'id="auth-password-reset-close"',
  'aria-live="polite"',
  'autocomplete="email"'
])assert.ok(html.includes(required),"Missing accessible recovery UI: "+required);
assert.doesNotMatch(source,/getPasswordResetContinueUrl|passwordReset.*complete|history\.replaceState/);
assert.match(source,/sendPasswordResetEmail\(email\)/);
assert.doesNotMatch(source,/fetchSignInMethodsForEmail|generatePasswordResetLink|confirmPasswordReset\(/);

function harness({email="",error=null}={}){
  const events=new Map(),mail=[],store=new Map(),timers=[];
  class Node {
    constructor(id,value=""){this.id=id;this.value=value;this.textContent="";this.disabled=false;
      this.open=false;this.handlers={};this.focused=false;this.resetCalls=0;}
    addEventListener(event,fn){this.handlers[event]=fn;}
    click(){assert.equal(typeof this.handlers.click,"function");return this.handlers.click({preventDefault(){}});}
    submit(){assert.equal(typeof this.handlers.submit,"function");
      return this.handlers.submit({preventDefault(){}});}
    showModal(){this.open=true;}
    close(){this.open=false;}
    reset(){this.resetCalls++;}
    focus(){this.focused=true;}
  }
  const nodes={};
  for(const name of [
    "first-password-dialog","first-password-form","first-password-input",
    "first-password-confirm-input","first-password-feedback","first-password-save-btn",
    "auth-email-input","auth-email-feedback","auth-request-password-btn",
    "auth-password-reset-dialog","auth-password-reset-form",
    "auth-password-reset-email","auth-password-reset-feedback",
    "auth-password-reset-submit","auth-password-reset-close"
  ])nodes[name]=new Node(name);
  nodes["auth-email-input"].value=email;
  const fakeAuth={
    languageCode:"en",
    onAuthStateChanged(fn){events.set("auth-change",fn);},
    async sendPasswordResetEmail(...args){
      mail.push(args);
      if(error)throw error;
    }
  };
  const firebase={auth:()=>fakeAuth,firestore:()=>({}),apps:[{}]};
  const window={
    firebase,
    localStorage:{
      getItem:key=>store.get(key)??null,
      setItem:(key,value)=>store.set(key,value)
    },
    setTimeout:(fn)=>{timers.push(fn);fn();return 1;}
  };
  const document={
    readyState:"complete",
    getElementById:name=>nodes[name]||null,
    documentElement:{classList:{remove(){}}}
  };
  vm.runInNewContext(source,{firebase,window,document,Date,Math,Number,Array,Object,Set,String,Promise,console:{error(){}}},{filename:"first-login-password.js"});
  return {nodes,mail,store,fakeAuth,events,click:()=>nodes["auth-request-password-btn"].click(),
    submit:()=>nodes["auth-password-reset-form"].submit(),close:()=>nodes["auth-password-reset-close"].click()};
}
(async()=>{
  const normal=harness();
  normal.click();
  assert.equal(normal.nodes["auth-password-reset-dialog"].open,true);
  assert.equal(normal.nodes["auth-password-reset-email"].focused,true);
  assert.equal(normal.mail.length,0);
  await normal.submit();
  assert.match(normal.nodes["auth-password-reset-feedback"].textContent,/indirizzo email valido/);
  assert.equal(normal.mail.length,0);

  const successful=harness({email:" FIRST.Last@Example.Test "});
  successful.click();
  assert.equal(successful.nodes["auth-password-reset-email"].value,"FIRST.Last@Example.Test");
  // The only input to the Firebase API is normalized email. Neither new
  // password nor app URL is sent or even accepted.
  await successful.submit();
  assert.equal(successful.mail.length,1);
  assert.deepEqual(successful.mail[0],["first.last@example.test"]);
  assert.equal(successful.fakeAuth.languageCode,"it");
  assert.equal(successful.nodes["auth-password-reset-submit"].disabled,false);
  assert.match(successful.nodes["auth-password-reset-feedback"].textContent,/Se l'indirizzo è associato/);
  assert.match(successful.nodes["auth-password-reset-feedback"].textContent,/torna qui per accedere/);
  assert.doesNotMatch(successful.nodes["auth-password-reset-feedback"].textContent,/Password aggiornata/);
  await successful.submit();
  assert.equal(successful.mail.length,1,"5-minute local duplicate throttling");
  assert.match(successful.nodes["auth-password-reset-feedback"].textContent,/Attendi circa/);
  await successful.close();
  assert.equal(successful.nodes["auth-password-reset-dialog"].open,false);
  assert.equal(successful.nodes["auth-password-reset-form"].resetCalls,1);

  const unknown=harness({email:"unknown@example.test",error:{code:"auth/user-not-found"}});
  unknown.click();await unknown.submit();
  assert.match(unknown.nodes["auth-password-reset-feedback"].textContent,/Se l'indirizzo è associato/);
  assert.doesNotMatch(unknown.nodes["auth-password-reset-feedback"].textContent,/non esiste|user-not-found/i);
  assert.equal(unknown.mail.length,1);
  assert.ok(unknown.store.size>0,"Unknown email must receive same local cooldown");
  const wrongProvider=harness({email:"google-only@example.test",error:{code:"auth/invalid-credential"}});
  wrongProvider.click();await wrongProvider.submit();
  assert.match(wrongProvider.nodes["auth-password-reset-feedback"].textContent,/Se l'indirizzo è associato/);

  const quota=harness({error:{code:"auth/too-many-requests"}});
  quota.click();quota.nodes["auth-password-reset-email"].value="somebody@example.test";
  await quota.submit();
  assert.match(quota.nodes["auth-password-reset-feedback"].textContent,/troppe richieste/i);
  assert.ok(quota.store.size>0);
  const network=harness({error:{code:"auth/network-request-failed"}});
  network.click();network.nodes["auth-password-reset-email"].value="somebody@example.test";
  await network.submit();
  assert.match(network.nodes["auth-password-reset-feedback"].textContent,/Connessione non disponibile/);
  assert.equal(network.store.size,0,"Network errors can be retried locally");

  const invalid=harness({email:"x@bad"});
  invalid.click();await invalid.submit();
  assert.equal(invalid.mail.length,0);
  const tooLong=harness({email:"x".repeat(260)+"@example.test"});
  tooLong.click();await tooLong.submit();
  assert.equal(tooLong.mail.length,0);

  // Signing in with Google/Apple is a separate authentication mechanism.
  assert.match(html,/Usi Google o Apple\?/);
  assert.ok(!source.includes("sendPasswordResetEmail(email, {"));
  console.log("PASS: self-service dialog, email validation, Firebase one-time link, non-enumeration, 5-min cooldown, no false completion, network retry, native/browser-neutral redirectless reset.");
})().catch(e=>{console.error(e);process.exitCode=1;});

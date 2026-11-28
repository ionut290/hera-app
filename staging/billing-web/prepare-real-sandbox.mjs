// Phase 1 of the opt-in hosted Stripe Sandbox test.
// Creates only synthetic emulator data and one real Stripe TEST Checkout Session.
import assert from 'node:assert/strict';
import {randomUUID} from 'node:crypto';
import {appendFileSync,chmodSync,writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {initializeApp,deleteApp} from 'firebase-admin/app';
import {getAuth} from 'firebase-admin/auth';
import {getFirestore} from 'firebase-admin/firestore';
import {TRIAL_LIMITS,initialTrialUsage} from '../trial-plan.mjs';
import {SANDBOX_PRICE_IDS,auditSandboxCatalog} from '../stripe-sandbox-catalog.mjs';
import {createSandboxStripeProvider,createSandboxCheckout} from '../stripe-sandbox-checkout.mjs';
const app=initializeApp({projectId:'varga-cantieri-v2-staging'});
const auth=getAuth(app),db=getFirestore(app);
let checkout,token,tenantId,created=false;
try{
 const provider=createSandboxStripeProvider({app,secretKey:process.env.VARGA_STRIPE_SANDBOX_CHECKOUT_KEY});
 await auditSandboxCatalog(provider);
 checkout=createSandboxCheckout({app,provider});
 const uid='sandbox-'+randomUUID(),email=uid+'@example.test',password=randomUUID()+'Aa42!';
 tenantId='demo-'+uid;
 await auth.createUser({uid,email,password,emailVerified:true});
 await db.doc('tenants/'+tenantId).set({name:'Varga Sandbox QA',demo:true,status:'active',plan:'trial',trial:true,trialDays:30,trialStartedAt:new Date(),limits:{...TRIAL_LIMITS},usage:initialTrialUsage(),ownerUid:uid});
 await db.doc('tenants/'+tenantId+'/members/'+uid).set({role:'admin',status:'active'});
 await db.doc('users/'+uid+'/memberships/'+tenantId).set({tenantId,role:'admin',status:'active'});
 const login=await fetch('http://127.0.0.1:9099/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=fake-api-key',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email,password,returnSecureToken:true})});
 assert.equal(login.ok,true);token=(await login.json()).idToken;
 await checkout.requestPlan({token,tenantId,tier:'starter'});
 const session=await checkout.createCheckout({token,tenantId,tier:'starter',acceptTerms:true});
 assert.match(session.sessionId,/^cs_test_/);assert.equal(session.testMode,true);
 const remote=await provider.session(session.sessionId);
 assert.equal(remote.livemode,false);assert.equal(remote.status,'open');assert.equal(remote.payment_status,'unpaid');
 assert.equal((await db.doc('tenants/'+tenantId).get()).data().plan,'trial');
 created=true;
 const statePath=join(process.env.RUNNER_TEMP,'varga-sandbox-checkout-state.json');
 writeFileSync(statePath,JSON.stringify({email,password,tenantId,sessionId:session.sessionId}));
 chmodSync(statePath,0o600);
 const summaryFile=process.env.GITHUB_STEP_SUMMARY;
 if(!summaryFile)throw new Error('SUMMARY_UNAVAILABLE');
 appendFileSync(summaryFile,'## Stripe Sandbox Checkout manuale\n\nApri [il Checkout Starter Sandbox]('+session.url+') e completa hCaptcha e la carta di prova Stripe.\n\nLa sessione è valida per questo test. Dopo il pagamento, GitHub verificherà automaticamente la fattura e l’entitlement nell’emulatore.\n');
 console.log('REAL SANDBOX CHECKOUT CREATED: test-only Starter session, 1990 EUR cents, no entitlement granted.');
 console.log('MANUAL CHECKOUT READY: link published in GitHub Actions summary; sensitive URL omitted from logs.');
}catch(e){
 console.error('REAL STRIPE SANDBOX PREPARATION FAILED; no sensitive provider details logged.');
 process.exitCode=1;
 if(created&&checkout&&token&&tenantId){try{await checkout.cancelCheckout({token,tenantId});console.log('Verified unpaid test checkout cleared.');}catch{console.error('Uncertain session retained for Sandbox reconciliation.');}}
}finally{await deleteApp(app);}

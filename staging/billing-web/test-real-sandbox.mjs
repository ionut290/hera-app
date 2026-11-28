// Opt-in external Sandbox integration test. NEVER use live credentials/cards or real user data.
// No screenshots, traces, URLs, credentials or provider error objects in logs.
import assert from 'node:assert/strict';
import {randomUUID} from 'node:crypto';
import {chromium} from 'playwright';
import {initializeApp,deleteApp} from 'firebase-admin/app';
import {getAuth} from 'firebase-admin/auth';
import {getFirestore} from 'firebase-admin/firestore';
import {TRIAL_LIMITS,initialTrialUsage} from '../trial-plan.mjs';
import {SUBSCRIPTION_PLANS} from '../subscription-plans.mjs';
import {SANDBOX_PRICE_IDS,auditSandboxCatalog} from '../stripe-sandbox-catalog.mjs';
import {createSandboxStripeProvider,createSandboxCheckout} from '../stripe-sandbox-checkout.mjs';
import {createStripeTestBilling} from '../stripe-test-billing.mjs';
const app=initializeApp({projectId:'varga-cantieri-v2-staging'}),db=getFirestore(app),auth=getAuth(app);
let browser,checkout,token,tenantId,stage='initialization',completed=false;
try{
 const provider=createSandboxStripeProvider({app,secretKey:process.env.VARGA_STRIPE_SANDBOX_CHECKOUT_KEY});
 await auditSandboxCatalog(provider);
 const billing=createStripeTestBilling({app,webhookSecret:'whsec_unused_external_reconciliation_fixture',priceIds:SANDBOX_PRICE_IDS,authoritativeProvider:provider});
 checkout=createSandboxCheckout({app,provider});
 const uid='sandbox-'+randomUUID(),email=uid+'@example.test',password=randomUUID()+'Aa42!';
 tenantId='demo-'+uid;
 await auth.createUser({uid,email,password,emailVerified:true});
 await db.doc('tenants/'+tenantId).set({name:'Varga Sandbox QA',demo:true,status:'active',plan:'trial',trial:true,trialDays:30,trialStartedAt:new Date(),limits:{...TRIAL_LIMITS},usage:initialTrialUsage(),ownerUid:uid});
 await db.doc('tenants/'+tenantId+'/members/'+uid).set({role:'admin',status:'active'});
 await db.doc('users/'+uid+'/memberships/'+tenantId).set({tenantId,role:'admin',status:'active'});
 const login=await fetch('http://127.0.0.1:9099/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=fake-api-key',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email,password,returnSecureToken:true})});
 assert.equal(login.ok,true);token=(await login.json()).idToken;
 stage='checkout-creation';
 await checkout.requestPlan({token,tenantId,tier:'starter'});
 const session=await checkout.createCheckout({token,tenantId,tier:'starter',acceptTerms:true});
 assert.match(session.sessionId,/^cs_test_/);assert.equal(session.testMode,true);
 const remote=await provider.session(session.sessionId);assert.equal(remote.livemode,false);assert.equal(remote.status,'open');assert.equal(remote.payment_status,'unpaid');
 assert.equal((await db.doc('tenants/'+tenantId).get()).data().plan,'trial');
 console.log('REAL SANDBOX CHECKOUT CREATED: test-only Starter session, 1990 EUR cents, no entitlement granted.');
 stage='hosted-checkout-browser';
 browser=await chromium.launch({headless:true});
 const page=await browser.newPage({locale:'en-US'});page.setDefaultTimeout(20000);
 await page.goto(session.url,{waitUntil:'domcontentloaded',timeout:45000});
 assert.equal(new URL(page.url()).hostname,'checkout.stripe.com');
 // Stripe's publicly documented test card; never substitute a real card.
 const fill=async(sel,value,required=true)=>{const el=page.locator(sel).first();if(required)await el.waitFor({state:'visible'});if(await el.isVisible())await el.fill(value);};
 await fill('#email, input[name=email]',email);
 await fill('#cardNumber, input[name=cardNumber]','4242424242424242');
 await fill('#cardExpiry, input[name=cardExpiry]','1230');
 await fill('#cardCvc, input[name=cardCvc]','123');
 await fill('#billingName, input[name=billingName]','Varga Sandbox QA',false);
 const country=page.locator('#billingCountry, select[name=billingCountry]').first();if(await country.isVisible())await country.selectOption('IT');
 const manual=page.getByText(/enter address manually/i).first();if(await manual.isVisible())await manual.click();
 await fill('#billingAddressLine1, input[name=billingAddressLine1]','Via Esempio 1',false);
 await fill('#billingLocality, input[name=billingLocality]','Bologna',false);
 await fill('#billingPostalCode, input[name=billingPostalCode]','40100',false);
 const region=page.locator('#billingAdministrativeArea, select[name=billingAdministrativeArea]').first();if(await region.isVisible()){const options=await region.locator('option').evaluateAll(ns=>ns.map(n=>({v:n.value,t:n.textContent})));const b=options.find(o=>o.v==='BO'||o.t==='Bologna');if(b)await region.selectOption(b.v);}
 const submit=page.locator('button[type=submit]').filter({hasText:/subscribe|pay/i}).first();
 await submit.waitFor({state:'visible'});await submit.click();
 stage='provider-payment-confirmation';
 // Poll the authoritative API, not the redirect URL. Never submit twice.
 for(let i=0;i<30;i++){const fresh=await provider.session(session.sessionId);if(fresh.status==='complete'&&fresh.payment_status==='paid'){completed=true;break;}await new Promise(r=>setTimeout(r,2000));}
 if(!completed)throw Error('PAYMENT_NOT_CONFIRMED');
 stage='emulator-entitlement';
 const bound=await checkout.finalizeCheckout({token,tenantId});
 assert.equal(bound.status,'bound-awaiting-invoice');
 assert.equal((await db.doc('tenants/'+tenantId).get()).data().plan,'trial');
 assert.equal((await billing.reconcileInitialInvoice(bound.subscriptionId)).status,'applied');
 const paid=(await db.doc('tenants/'+tenantId).get()).data();
 assert.equal(paid.plan,'full');assert.equal(paid.billing.amountCents,1990);assert.equal(paid.billing.periods,1);assert.deepEqual(paid.limits,SUBSCRIPTION_PLANS.starter.limits);
 assert.equal((await billing.reconcileInitialInvoice(bound.subscriptionId)).status,'duplicate-event');
 assert.equal((await db.doc('tenants/'+tenantId).collection('billingReceipts').get()).size,1);
 console.log('REAL STRIPE SANDBOX PAYMENT PASS: hosted test card, paid invoice and succeeded PaymentIntent verified, Starter entitlement applied once to emulator only.');
 console.log('A synthetic test subscription remains in Stripe Sandbox; cancel it in the dashboard after verification. No live transaction.');
}catch(e){
 console.error('REAL STRIPE SANDBOX TEST FAILED at '+stage+'; no sensitive provider details logged.');
 process.exitCode=1;
 if(checkout&&token&&tenantId&&!completed){try{await checkout.cancelCheckout({token,tenantId});console.log('Verified unpaid test checkout cleared.');}catch{console.error('Uncertain or completed session retained for manual Sandbox reconciliation.');}}
}finally{if(browser)await browser.close();await deleteApp(app);}

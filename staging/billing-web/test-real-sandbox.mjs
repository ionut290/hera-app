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
 const page=await browser.newPage({locale:'en-US'});page.setDefaultTimeout(45000);
 await page.goto(session.url,{waitUntil:'domcontentloaded',timeout:60000});
 assert.equal(new URL(page.url()).hostname,'checkout.stripe.com');
 // Stripe Checkout hosts payment inputs in cross-origin frames. Search the page
 // and its frames, but keep all entered data synthetic and test-only.
 const describeCheckoutFields=async()=>{
  const summary=[];
  for(const [index,frame] of page.frames().entries()){
   let inputs=[];try{inputs=await frame.locator('input').evaluateAll(ns=>ns.slice(0,20).map(n=>({type:n.type,name:n.getAttribute('name'),id:n.id,autocomplete:n.getAttribute('autocomplete'),ariaLabel:n.getAttribute('aria-label'),placeholder:n.getAttribute('placeholder')})));}catch{}
   let host='unknown';try{host=new URL(frame.url()).hostname;}catch{}
   summary.push({frame:index,host,inputs});
  }
  console.log('CHECKOUT_FIELDS '+JSON.stringify(summary));
 };
 const fill=async(sel,value,required=true)=>{
  const selectors=sel.split(',').map(s=>s.trim()).filter(Boolean),deadline=Date.now()+(required?45000:6000);
  do{
   for(const context of [page,...page.frames().filter(f=>f!==page.mainFrame())]){
    for(const selector of selectors){const el=context.locator(selector).first();try{if(await el.isVisible()){await el.fill(value,{timeout:5000});return true;}}catch{}}
   }
   await page.waitForTimeout(400);
  }while(Date.now()<deadline);
  if(!required)return false;
  await describeCheckoutFields();
  throw new Error('CHECKOUT_FIELD_UNAVAILABLE');
 };
 await describeCheckoutFields();
 stage='hosted-checkout-email';
 await fill('#email, input[name=email], input[type=email], input[autocomplete=email], input[placeholder*="email" i]',email,false);
 stage='hosted-checkout-card-number';
 await fill('#cardNumber, input[name=cardNumber], input[name=cardnumber], input[autocomplete=cc-number], input[aria-label*="card number" i], input[placeholder*="1234" i]','4242424242424242');
 stage='hosted-checkout-card-expiry';
 await fill('#cardExpiry, input[name=cardExpiry], input[autocomplete=cc-exp], input[aria-label*="expiration" i]','1230');
 stage='hosted-checkout-card-cvc';
 await fill('#cardCvc, input[name=cardCvc], input[autocomplete=cc-csc], input[aria-label*="security code" i]','123');
 stage='hosted-checkout-billing-name';
 await fill('#billingName, input[name=billingName], input[autocomplete=cc-name]','Varga Sandbox QA',false);
 stage='hosted-checkout-country';
 const country=page.locator('#billingCountry, select[name=billingCountry]').first();
 if(await country.isVisible()){const options=await country.locator('option').evaluateAll(ns=>ns.map(n=>({v:n.value,t:n.textContent})));const it=options.find(o=>o.v==='IT'||/italy|italia/i.test(o.t||''));if(it)await country.selectOption(it.v);}
 stage='hosted-checkout-address-expand';
 const manual=page.getByText(/enter address manually/i).first();if(await manual.isVisible())await manual.click();
 stage='hosted-checkout-address-line';
 await fill('#billingAddressLine1, input[name=billingAddressLine1], input[autocomplete=address-line1]','Via Esempio 1',false);
 stage='hosted-checkout-city';
 await fill('#billingLocality, input[name=billingLocality], input[autocomplete=address-level2]','Bologna',false);
 stage='hosted-checkout-postal-code';
 await fill('#billingPostalCode, input[name=billingPostalCode], input[autocomplete=postal-code]','40100',false);
 stage='hosted-checkout-region';
 const region=page.locator('#billingAdministrativeArea, select[name=billingAdministrativeArea]').first();
 if(await region.isVisible()){const options=await region.locator('option').evaluateAll(ns=>ns.map(n=>({v:n.value,t:n.textContent})));const b=options.find(o=>o.v==='BO'||o.t==='Bologna');if(b)await region.selectOption(b.v);}
 stage='hosted-checkout-submit';
 const visibleButtons=await page.locator('button').evaluateAll(ns=>ns.filter(n=>!!(n.offsetWidth||n.offsetHeight)).map(n=>({text:(n.innerText||'').trim().slice(0,80),type:n.getAttribute('type'),ariaLabel:n.getAttribute('aria-label'),disabled:!!n.disabled})).slice(0,20));
 console.log('CHECKOUT_SUBMIT_OPTIONS '+JSON.stringify(visibleButtons));
 const submitSpecific=page.locator('button[type=submit]').filter({hasText:/subscribe|pay/i}).first();
 const byRole=page.getByRole('button',{name:/subscribe|pay|complete|confirm|continue/i}).last();
 const visibleButtons=page.locator('button').filter({visible:true});
 const visibleButtonCount=await visibleButtons.count();
 const submit=await submitSpecific.count()?submitSpecific:await byRole.count()?byRole:visibleButtonCount===1?visibleButtons.first():null;
 if(!submit)throw new Error('CHECKOUT_SUBMIT_AMBIGUOUS');
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

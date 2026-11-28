// Phase 2 of the opt-in hosted Stripe Sandbox test.
// Polls the authoritative Stripe TEST API, then grants entitlement only in the emulator.
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {join} from 'node:path';
import {initializeApp,deleteApp} from 'firebase-admin/app';
import {getAuth} from 'firebase-admin/auth';
import {getFirestore} from 'firebase-admin/firestore';
import {SUBSCRIPTION_PLANS} from '../subscription-plans.mjs';
import {SANDBOX_PRICE_IDS} from '../stripe-sandbox-catalog.mjs';
import {createSandboxStripeProvider,createSandboxCheckout} from '../stripe-sandbox-checkout.mjs';
import {createStripeTestBilling} from '../stripe-test-billing.mjs';
const app=initializeApp({projectId:'varga-cantieri-v2-staging'}),db=getFirestore(app),auth=getAuth(app);
let checkout,token,tenantId,completed=false;
try{
 const state=JSON.parse(readFileSync(join(process.env.RUNNER_TEMP,'varga-sandbox-checkout-state.json'),'utf8'));
 tenantId=state.tenantId;
 const provider=createSandboxStripeProvider({app,secretKey:process.env.VARGA_STRIPE_SANDBOX_CHECKOUT_KEY});
 checkout=createSandboxCheckout({app,provider});
 const billing=createStripeTestBilling({app,webhookSecret:'whsec_unused_external_reconciliation_fixture',priceIds:SANDBOX_PRICE_IDS,authoritativeProvider:provider});
 const login=await fetch('http://127.0.0.1:9099/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=fake-api-key',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email:state.email,password:state.password,returnSecureToken:true})});
 assert.equal(login.ok,true);token=(await login.json()).idToken;
 let session;
 for(let i=0;i<1050;i++){
  session=await provider.session(state.sessionId);
  if(session.status==='complete'&&session.payment_status==='paid'){completed=true;break;}
  if(session.status==='expired')break;
  await new Promise(r=>setTimeout(r,2000));
 }
 if(!completed)throw new Error('MANUAL_PAYMENT_NOT_CONFIRMED');
 const bound=await checkout.finalizeCheckout({token,tenantId});
 assert.equal(bound.status,'bound-awaiting-invoice');
 assert.equal((await db.doc('tenants/'+tenantId).get()).data().plan,'trial');
 assert.equal((await billing.reconcileInitialInvoice(bound.subscriptionId)).status,'applied');
 const paid=(await db.doc('tenants/'+tenantId).get()).data();
 assert.equal(paid.plan,'full');assert.equal(paid.billing.amountCents,1990);assert.equal(paid.billing.periods,1);assert.deepEqual(paid.limits,SUBSCRIPTION_PLANS.starter.limits);
 assert.equal((await billing.reconcileInitialInvoice(bound.subscriptionId)).status,'duplicate-event');
 assert.equal((await db.doc('tenants/'+tenantId).collection('billingReceipts').get()).size,1);
 console.log('REAL STRIPE SANDBOX PAYMENT PASS: paid invoice and succeeded PaymentIntent verified; Starter entitlement applied once to emulator only.');
 console.log('A synthetic test subscription remains in Stripe Sandbox; cancel it in the dashboard after verification. No live transaction.');
}catch(e){
 console.error('REAL STRIPE SANDBOX VERIFICATION FAILED; no sensitive provider details logged.');
 process.exitCode=1;
 if(checkout&&token&&tenantId&&!completed){try{await checkout.cancelCheckout({token,tenantId});console.log('Verified unpaid test checkout cleared.');}catch{console.error('Uncertain or completed session retained for manual Sandbox reconciliation.');}}
}finally{await deleteApp(app);}

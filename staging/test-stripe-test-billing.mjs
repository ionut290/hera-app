import assert from 'node:assert/strict';
import Stripe from 'stripe';
import {initializeApp,deleteApp} from 'firebase-admin/app';
import {getFirestore} from 'firebase-admin/firestore';
import {createStripeTestBilling} from './stripe-test-billing.mjs';
import {SUBSCRIPTION_PLANS,periodEnd} from './subscription-plans.mjs';
import {TRIAL_LIMITS,initialTrialUsage} from './trial-plan.mjs';
assert.equal(process.env.FIRESTORE_EMULATOR_HOST,'127.0.0.1:8080');
const app=initializeApp({projectId:'varga-cantieri-v2-staging'}),db=getFirestore(app);
const secret='whsec_local_fixture_only_not_a_real_secret',priceIds={starter:'price_fixture_starter',professional:'price_fixture_professional',business:'price_fixture_business'};
const service=createStripeTestBilling({app,webhookSecret:secret,priceIds});
const stripe=new Stripe('sk_test_LOCAL_FIXTURE_ONLY',{apiVersion:'2024-10-28.acacia'});
const A='demo-stripe-a',B='demo-stripe-b',sub='sub_fixture_a',customer='cus_fixture_a';
const span=30*86400,start=Math.floor(Date.now()/1000)-span+120,end=start+span;
const err=(code,fn)=>assert.rejects(fn,e=>e.code===code);
function event(id,invoice,kind='invoice.paid'){return {id:'evt_fixture_'+id,object:'event',api_version:'2024-10-28.acacia',livemode:false,type:kind,created:Math.floor(Date.now()/1000),data:{object:invoice}};}
function invoice(id,from=start,reason='subscription_create'){return {id:'in_fixture_'+id,object:'invoice',livemode:false,subscription:sub,customer,status:'paid',paid:true,paid_out_of_band:false,payment_intent:'pi_fixture_'+id,currency:'eur',amount_paid:3990,amount_due:3990,amount_remaining:0,collection_method:'charge_automatically',billing_reason:reason,lines:{has_more:false,data:[{type:'subscription',proration:false,quantity:1,amount:3990,currency:'eur',price:{id:priceIds.professional,recurring:{interval:'day',interval_count:30}},period:{start:from,end:from+span}}]}};}
function signed(e){const payload=JSON.stringify(e);return [Buffer.from(payload),stripe.webhooks.generateTestHeaderString({payload,secret})];}
const send=e=>service.handleWebhook(...signed(e));
try{
 const trial={demo:true,status:'active',plan:'trial',trial:true,trialDays:30,trialStartedAt:new Date(),limits:{...TRIAL_LIMITS},usage:initialTrialUsage(),requestedTier:'professional',upgradeRequested:true};
 await db.doc('tenants/'+A).set(trial);await db.doc('tenants/'+B).set(trial);
 await err('UNBOUND_SUBSCRIPTION',()=>send(event('unbound',invoice('unbound'))));
 await service.bindTestSubscription({tenantId:A,subscriptionId:sub,customerId:customer,tier:'professional'});
 await err('BINDING_CONFLICT',()=>service.bindTestSubscription({tenantId:B,subscriptionId:sub,customerId:customer,tier:'professional'}));
 await err('BINDING_CONFLICT',()=>service.bindTestSubscription({tenantId:A,subscriptionId:'sub_fixture_other',customerId:customer,tier:'professional'}));
 const e=event('first',invoice('first')),args=signed(e);
 await err('INVALID_SIGNATURE',()=>service.handleWebhook(args[0],args[1]+'bad'));
 await err('INVALID_SIGNATURE',()=>service.handleWebhook(Buffer.from(JSON.stringify({...e,created:0})),args[1]));
 await err('INVALID_EVENT',()=>send({...e,livemode:true}));
 await err('INVOICE_MISMATCH',()=>send(event('wrong_amount',{...invoice('wrong_amount'),amount_paid:1990})));
 await err('INVOICE_MISMATCH',()=>send(event('out_of_band',{...invoice('out_of_band'),paid_out_of_band:true})));
 await err('UNBOUND_SUBSCRIPTION',()=>send(event('wrong_customer',{...invoice('wrong_customer'),customer:'cus_fixture_other'})));
 assert.equal((await db.doc('tenants/'+A).get()).data().plan,'trial');
 const both=await Promise.all([send(e),send(e)]);assert.equal(both.filter(x=>x.status==='applied').length,1);assert.equal(both.filter(x=>x.status==='duplicate-event').length,1);
 let t=(await db.doc('tenants/'+A).get()).data();assert.equal(t.plan,'full');assert.equal(t.billing.source,'stripe-test');assert.equal(t.billing.periods,1);assert.equal(periodEnd(t),end*1000);assert.deepEqual(t.limits,SUBSCRIPTION_PLANS.professional.limits);assert.deepEqual(t.usage,trial.usage);
 assert.equal((await send(event('same_invoice',invoice('first')))).status,'duplicate-invoice');
 await err('PERIOD_ALREADY_PAID',()=>send(event('different_invoice',invoice('different_invoice'))));
 assert.equal((await send(event('failure',{...invoice('failed',end,'subscription_cycle'),paid:false,status:'open',amount_paid:0},'invoice.payment_failed'))).status,'payment-failed');
 assert.equal((await db.doc('tenants/'+A).get()).data().billing.periods,1);
 await err('INVOICE_MISMATCH',()=>send(event('future',invoice('future',end+span,'subscription_cycle'))));
 assert.equal((await send(event('renew',invoice('renew',end,'subscription_cycle')))).status,'applied');
 t=(await db.doc('tenants/'+A).get()).data();assert.equal(t.billing.periods,2);assert.equal(periodEnd(t),(end+span)*1000);assert.deepEqual(t.usage,trial.usage);
 assert.equal((await db.doc('tenants/'+A).collection('billingReceipts').get()).size,2);
 assert.equal((await db.collection('stripeTestInvoices').get()).size,2);
 assert.equal((await db.doc('tenants/'+B).get()).data().plan,'trial');
 assert.equal((await send(event('cancel',{id:sub,customer,livemode:false,status:'canceled'},'customer.subscription.deleted'))).status,'canceled');
 assert.equal((await db.doc('stripeTestSubscriptions/'+sub).get()).data().status,'canceled');
 assert.equal(periodEnd((await db.doc('tenants/'+A).get()).data()),(end+span)*1000);
 // A canceled binding must not extend an existing paid period, even for an otherwise valid invoice.
 await err('PERIOD_ALREADY_PAID',()=>send(event('cancel_duplicate_period',invoice('cancel_duplicate_period',end,'subscription_cycle'))));
 console.log('STRIPE TEST PASS: official SDK signature, live events rejected, verified binding, exact paid amount, duplicate event/invoice/period protection, automatic renewal, failure without entitlement, cancellation retaining paid period, company isolation.');
}finally{await deleteApp(app);}

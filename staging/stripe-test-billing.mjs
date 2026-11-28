// Emulator-only Stripe adapter. No Stripe API calls, no Checkout URL, no live payments.
import Stripe from 'stripe';
import {getFirestore,FieldValue,Timestamp} from 'firebase-admin/firestore';
import {SUBSCRIPTION_PLANS,BILLING_PERIOD_DAYS} from './subscription-plans.mjs';
const PROJECT='varga-cantieri-v2-staging',VERSION='2024-10-28.acacia';
const SECONDS=BILLING_PERIOD_DAYS*86400;
const id=(v,p)=>typeof v==='string'&&v.startsWith(p)&&/^[A-Za-z0-9_]{5,100}$/.test(v);
const tenantIdOk=v=>typeof v==='string'&&/^demo-[a-z0-9-]{3,59}$/.test(v);
const fail=code=>{throw Object.assign(new Error(code),{code});};
const plans=Object.keys(SUBSCRIPTION_PLANS);
const usageFits=(u,l)=>Object.entries(l).every(([k,v])=>Number.isSafeInteger(u?.[k])&&u[k]>=0&&u[k]<=v);
export function createStripeTestBilling({app,webhookSecret,priceIds,authoritativeProvider=null}){
 if(app?.options?.projectId!==PROJECT||process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8080')fail('EMULATORS_REQUIRED');
 if(typeof webhookSecret!=='string'||!webhookSecret.startsWith('whsec_')||webhookSecret.length<20)fail('INVALID_CONFIG');
 if(!plans.every(t=>id(priceIds?.[t],'price_'))||new Set(plans.map(t=>priceIds[t])).size!==plans.length)fail('INVALID_CONFIG');
 const db=getFirestore(app),stripe=new Stripe('sk_test_LOCAL_FIXTURE_ONLY',{apiVersion:VERSION,maxNetworkRetries:0});
 const stamp=()=>FieldValue.serverTimestamp();
 // Trusted backend provisioning only. No public method or user-supplied tenant binding.
 async function bindTestSubscription({tenantId,subscriptionId,customerId,tier}){
  if(!tenantIdOk(tenantId)||!id(subscriptionId,'sub_')||!id(customerId,'cus_')||!plans.includes(tier))fail('INVALID_BINDING');
  const tenant=db.doc('tenants/'+tenantId),binding=db.doc('stripeTestSubscriptions/'+subscriptionId);
  return db.runTransaction(async tx=>{
   const [ts,bs]=await Promise.all([tx.get(tenant),tx.get(binding)]),t=ts.data(),b=bs.data();
   if(b){if(b.tenantId===tenantId&&b.customerId===customerId&&b.tier===tier&&t?.stripeTestSubscriptionId===subscriptionId)return {status:'already-bound'};fail('BINDING_CONFLICT');}
   if(!t||t.demo!==true||t.status!=='active'||t.plan!=='trial'||t.billing||t.stripeTestSubscriptionId||t.stripeSandboxCheckoutId||t.requestedTier!==tier||t.upgradeRequested!==true)fail('BINDING_CONFLICT');
   if(!usageFits(t.usage,SUBSCRIPTION_PLANS[tier].limits))fail('USAGE_EXCEEDS_PLAN');
   tx.create(binding,{tenantId,customerId,tier,subscriptionId,createdAt:stamp(),status:'pending'});
   tx.update(tenant,{stripeTestSubscriptionId:subscriptionId});
   return {status:'bound'};
  });
 }
 function readPaidInvoice(i,b,now){
  const p=SUBSCRIPTION_PLANS[b.tier],line=i.lines?.data?.[0],start=line?.period?.start,end=line?.period?.end;
  if(i.livemode!==false||i.status!=='paid'||i.paid!==true||i.paid_out_of_band!==false||!id(i.payment_intent,'pi_')||
   i.currency!=='eur'||i.amount_paid!==p.priceCents||i.amount_due!==p.priceCents||i.amount_remaining!==0||
   i.collection_method!=='charge_automatically'||!['subscription_create','subscription_cycle'].includes(i.billing_reason)||
   i.lines?.has_more!==false||i.lines?.data?.length!==1||line.type!=='subscription'||line.proration!==false||
   line.quantity!==1||line.amount!==p.priceCents||line.currency!=='eur'||line.price?.id!==priceIds[b.tier]||
   line.price?.recurring?.interval!=='day'||line.price?.recurring?.interval_count!==BILLING_PERIOD_DAYS||
   !Number.isSafeInteger(start)||!Number.isSafeInteger(end)||start<=0||end-start!==SECONDS||start>now+300)fail('INVOICE_MISMATCH');
  return {p,start,end};
 }
 async function handleWebhook(rawBody,signature){
  if(!Buffer.isBuffer(rawBody)||rawBody.length>1024*1024||typeof signature!=='string')fail('INVALID_SIGNATURE');
  let event;try{event=stripe.webhooks.constructEvent(rawBody,signature,webhookSecret,300);}catch{fail('INVALID_SIGNATURE');}
  return applyVerifiedEvent(event);
 }
 // Internal ledger entry point. Never accepts an event from an HTTP caller.
 async function applyVerifiedEvent(event){
  if(event.livemode!==false||event.api_version!==VERSION||event.account||!id(event.id,'evt_'))fail('INVALID_EVENT');
  if(!['invoice.paid','invoice.payment_failed','customer.subscription.deleted'].includes(event.type))return {status:'ignored'};
  const i=event.data?.object,subId=event.type==='customer.subscription.deleted'?i?.id:i?.subscription;
  if(!id(subId,'sub_')||!id(i?.customer,'cus_')||i.livemode!==false)fail('INVALID_EVENT');
  if(event.type!=='customer.subscription.deleted'&&!id(i.id,'in_'))fail('INVALID_EVENT');
  const binding=db.doc('stripeTestSubscriptions/'+subId),eventRef=db.doc('stripeTestEvents/'+event.id);
  return db.runTransaction(async tx=>{
   const [bs,es]=await Promise.all([tx.get(binding),tx.get(eventRef)]),b=bs.data();
   if(es.exists)return {status:'duplicate-event'};
   if(!b||!tenantIdOk(b.tenantId)||b.customerId!==i.customer||!plans.includes(b.tier))fail('UNBOUND_SUBSCRIPTION');
   const tenant=db.doc('tenants/'+b.tenantId),ts=await tx.get(tenant),t=ts.data();
   if(!t||t.demo!==true||t.status!=='active'||t.stripeTestSubscriptionId!==subId)fail('BINDING_CONFLICT');
   const audit={type:event.type,tenantId:b.tenantId,subscriptionId:subId,receivedAt:stamp(),testMode:true};
   if(event.type==='customer.subscription.deleted'){
    if(i.status!=='canceled')fail('INVALID_EVENT');
    tx.update(binding,{status:'canceled',canceledAt:stamp()});tx.create(eventRef,{...audit,outcome:'canceled'});return {status:'canceled'};
   }
   if(event.type==='invoice.payment_failed'){
    tx.create(eventRef,{...audit,invoiceId:i.id,outcome:'payment-failed'});return {status:'payment-failed'};
   }
   const {p,start,end}=readPaidInvoice(i,b,Math.floor(Date.now()/1000));
   const receipt=db.doc('stripeTestInvoices/'+i.id),period=db.doc('stripeTestSubscriptions/'+subId+'/periods/'+start),rs=await tx.get(receipt),ps=await tx.get(period);
   if(rs.exists){const d=rs.data();if(d.tenantId!==b.tenantId||d.subscriptionId!==subId||d.start!==start||d.end!==end)fail('INVOICE_CONFLICT');tx.create(eventRef,{...audit,invoiceId:i.id,outcome:'duplicate-invoice'});return {status:'duplicate-invoice'};}
   if(ps.exists)fail('PERIOD_ALREADY_PAID');
   if(b.status==='canceled')fail('SUBSCRIPTION_CANCELED');
   if(!usageFits(t.usage,p.limits))fail('USAGE_EXCEEDS_PLAN');
   let anchor=Timestamp.fromMillis(start*1000),periods=1;
   if(t.billing){
    if(t.plan!=='full'||t.billing.source!=='stripe-test'||t.billing.tier!==b.tier||!Number.isSafeInteger(t.billing.periods)||t.billing.periods<1||t.billing.periods>=120)fail('BILLING_CONFLICT');
    const previousEnd=t.billing.periodStart.toMillis()/1000+t.billing.periods*SECONDS;
    if(start!==previousEnd||i.billing_reason!=='subscription_cycle')fail('PERIOD_OUT_OF_ORDER');
    anchor=t.billing.periodStart;periods=t.billing.periods+1;
   }else if(t.plan!=='trial'||i.billing_reason!=='subscription_create'||end<=Math.floor(Date.now()/1000))fail('BILLING_CONFLICT');
   const billing={tier:b.tier,amountCents:p.priceCents,currency:'EUR',periodStart:anchor,periods,reference:i.id,recordedAt:stamp(),recordedBy:'stripe-test-service',source:'stripe-test'};
   const ledger={tenantId:b.tenantId,subscriptionId:subId,invoiceId:i.id,start,end,amountCents:p.priceCents,currency:'EUR',testMode:true,recordedAt:stamp()};
   tx.create(receipt,ledger);tx.create(period,ledger);tx.create(tenant.collection('billingReceipts').doc(i.id),billing);
   tx.update(tenant,{plan:'full',trial:false,limits:{...p.limits},billing,planUpdatedAt:stamp(),planUpdatedBy:'stripe-test-service',upgradeRequested:false,requestedTier:null});
   tx.update(binding,{status:'active',lastPaidInvoice:i.id,paidThrough:Timestamp.fromMillis(end*1000)});
   tx.create(eventRef,{...audit,invoiceId:i.id,outcome:'applied'});
   return {status:'applied',tier:b.tier,periods};
  });
 }
 async function reconcileInitialInvoice(subscriptionId){
  // The trusted server provider retrieves the invoice and succeeded PaymentIntent.
  // Caller supplies only a previously bound subscription identifier, not a payment.
  if(!id(subscriptionId,'sub_')||typeof authoritativeProvider?.initialPaidInvoice!=='function')fail('RECONCILIATION_UNAVAILABLE');
  const bound=(await db.doc('stripeTestSubscriptions/'+subscriptionId).get()).data();
  if(!bound||bound.status==='canceled')fail('UNBOUND_SUBSCRIPTION');
  const invoice=await authoritativeProvider.initialPaidInvoice(subscriptionId);
  if(!invoice||invoice.subscription!==subscriptionId||invoice.customer!==bound.customerId||invoice.billing_reason!=='subscription_create'||!id(invoice.id,'in_'))fail('INVOICE_MISMATCH');
  return applyVerifiedEvent({id:'evt_reconcile_'+invoice.id,api_version:VERSION,livemode:false,type:'invoice.paid',data:{object:invoice}});
 }
 return Object.freeze({bindTestSubscription,handleWebhook,reconcileInitialInvoice});
}
